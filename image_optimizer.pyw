import sys
import ctypes
import json
import time
import shutil
import zipfile
import subprocess
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Any

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QHBoxLayout, QMessageBox, QProgressBar, QSpinBox, QCheckBox, QGroupBox,
    QSlider, QListWidget, QListWidgetItem, QDialog, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont, QPalette, QColor, QIcon


SUPPORTED_EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")
OUTPUT_DIR_NAME = "webp_optimized"
ENCODER_NAME = "cwebp"
HISTORY_FILE = Path(__file__).with_suffix(".history.json")
MAX_HISTORY_ENTRIES = 20


@dataclass
class FolderBatch:
    folder: Path
    all_images: List[Path]
    convertible_images: List[Path]
    skipped_webp: List[Path]


@dataclass
class FolderSummary:
    folder: str
    converted: int
    skipped_existing: int
    errors: List[str]
    bytes_original: int
    bytes_converted: int
    archive_size: Optional[int]
    archive_path: Optional[str]
    duration_seconds: float


class Worker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal()
    summary_ready = pyqtSignal(object)

    def __init__(self, folders, quality, use_cbz, replace_originals, skip_webp):
        super().__init__()
        self.folders = folders
        self.quality = int(quality)
        self.use_cbz = use_cbz
        self.replace_originals = replace_originals
        self.skip_webp = skip_webp
        self._creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self._run_start = 0.0

    def run(self):
        run_start = time.perf_counter()
        folder_summaries: List[FolderSummary] = []
        processed_units = 0

        if not self._encoder_available():
            self.summary_ready.emit(
                self._build_run_summary(
                    folder_summaries, run_start, 0, processed_units, 0)
            )
            self.finished.emit()
            return

        batches, total_convertible, total_files = self._prepare_batches()
        total_work_units = total_files if total_files > 0 else total_convertible

        if total_convertible == 0:
            self.status.emit("No images to convert in selected folder(s).")
            self.summary_ready.emit(
                self._build_run_summary(
                    folder_summaries,
                    run_start,
                    total_work_units,
                    processed_units,
                    total_convertible
                )
            )
            self.finished.emit()
            return

        for batch in batches:
            if self.isInterruptionRequested():
                self.status.emit("Conversion cancelled.")
                self.summary_ready.emit(
                    self._build_run_summary(
                        folder_summaries,
                        run_start,
                        total_work_units,
                        processed_units,
                        total_convertible,
                        cancelled=True,
                    )
                )
                self.finished.emit()
                return

            folder_start = time.perf_counter()
            folder_errors: List[str] = []
            bytes_original = 0
            bytes_converted = 0
            archive_path: Optional[Path] = None
            archive_size: Optional[int] = None
            converted_count = 0

            self.status.emit(f"Processing folder '{batch.folder.name}'...")
            try:
                outdir = self._prepare_output_dir(batch.folder)
            except Exception as exc:
                error_msg = f"Unable to prepare output directory for {batch.folder.name}: {exc}"
                self.status.emit(error_msg)
                folder_errors.append(error_msg)
                folder_summary = FolderSummary(
                    folder=str(batch.folder),
                    converted=0,
                    skipped_existing=len(batch.skipped_webp),
                    errors=folder_errors,
                    bytes_original=0,
                    bytes_converted=0,
                    archive_size=None,
                    archive_path=None,
                    duration_seconds=time.perf_counter() - folder_start,
                )
                folder_summaries.append(folder_summary)
                continue

            if not batch.convertible_images and not self.replace_originals:
                message = f"No images to convert in {batch.folder.name}, skipping."
                self.status.emit(message)
                folder_errors.append(message)
                self._cleanup_dir(outdir)
                folder_summary = FolderSummary(
                    folder=str(batch.folder),
                    converted=0,
                    skipped_existing=len(batch.skipped_webp),
                    errors=folder_errors,
                    bytes_original=0,
                    bytes_converted=0,
                    archive_size=None,
                    archive_path=None,
                    duration_seconds=time.perf_counter() - folder_start,
                )
                folder_summaries.append(folder_summary)
                continue

            for source in batch.convertible_images:
                if self.isInterruptionRequested():
                    self.status.emit("Conversion cancelled.")
                    self._cleanup_dir(outdir)
                    self.summary_ready.emit(
                        self._build_run_summary(
                            folder_summaries,
                            run_start,
                            total_work_units,
                            processed_units,
                            total_convertible,
                            cancelled=True,
                        )
                    )
                    self.finished.emit()
                    return

                self.status.emit(
                    f"[{batch.folder.name}] Processing {source.name}...")
                success, original_size, converted_size, error_message = self._convert_image(
                    source, outdir)
                processed_units += 1
                self._emit_progress(processed_units, total_work_units)

                if success:
                    converted_count += 1
                    bytes_original += original_size
                    bytes_converted += converted_size
                elif error_message:
                    folder_errors.append(error_message)

            if self.replace_originals:
                errors = self._replace_originals(batch, outdir)
                folder_errors.extend(errors)
            else:
                archive_path, archive_size, archive_errors = self._create_archive(
                    batch, outdir)
                folder_errors.extend(archive_errors)

            skipped_units = len(batch.all_images) - \
                len(batch.convertible_images)
            if skipped_units > 0:
                processed_units += skipped_units
                self._emit_progress(processed_units, total_work_units)

            folder_summary = FolderSummary(
                folder=str(batch.folder),
                converted=converted_count,
                skipped_existing=len(batch.skipped_webp),
                errors=folder_errors,
                bytes_original=bytes_original,
                bytes_converted=bytes_converted,
                archive_size=archive_size,
                archive_path=str(archive_path) if archive_path else None,
                duration_seconds=time.perf_counter() - folder_start,
            )
            folder_summaries.append(folder_summary)

        self.status.emit("All done!")
        self.progress.emit(100)
        self.summary_ready.emit(
            self._build_run_summary(
                folder_summaries,
                run_start,
                total_work_units,
                processed_units,
                total_convertible,
            )
        )
        self.finished.emit()

    def _encoder_available(self) -> bool:
        if shutil.which(ENCODER_NAME):
            return True

        self.status.emit(
            f"Unable to find '{ENCODER_NAME}' executable. Please install it and make sure it is on PATH."
        )
        return False

    def _prepare_batches(self) -> tuple[List[FolderBatch], int, int]:
        batches: List[FolderBatch] = []
        total_convertible = 0
        total_files = 0

        for folder in self.folders:
            if not folder.exists():
                self.status.emit(
                    f"Folder '{folder}' does not exist, skipping.")
                continue

            all_images = [
                path for path in folder.iterdir()
                if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
            ]
            skipped_webp = [
                path for path in all_images if path.suffix.lower() == ".webp"
            ] if self.skip_webp else []
            convertible_images = [
                path for path in all_images
                if not (self.skip_webp and path.suffix.lower() == ".webp")
            ]

            batches.append(FolderBatch(folder, all_images,
                           convertible_images, skipped_webp))
            total_convertible += len(convertible_images)
            total_files += len(all_images)

        return batches, total_convertible, total_files

    def _prepare_output_dir(self, folder: Path) -> Path:
        outdir = folder / OUTPUT_DIR_NAME
        if outdir.exists():
            shutil.rmtree(outdir)
        outdir.mkdir()
        return outdir

    def _convert_image(self, source: Path, outdir: Path) -> tuple[bool, int, int, Optional[str]]:
        target = outdir / (source.stem + ".webp")
        cmd = self._build_cwebp_command(source, target)
        try:
            original_size = source.stat().st_size
        except OSError:
            original_size = 0
        error_message: Optional[str] = None

        try:
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=self._creationflags,
                check=False
            )
        except FileNotFoundError:
            self.status.emit(
                f"'{ENCODER_NAME}' executable not found while converting {source.name}.")
            error_message = f"{source.name}: encoder '{ENCODER_NAME}' not found."
            return False, original_size, 0, error_message
        except Exception as exc:
            self.status.emit(
                f"Unexpected error converting {source.name}: {exc}")
            error_message = f"{source.name}: unexpected error {exc}"
            return False, original_size, 0, error_message

        if result.returncode != 0:
            stderr = result.stderr.decode(errors="ignore") or "Unknown error."
            stderr = stderr.strip()
            self.status.emit(f"Error converting {source.name}: {stderr}")
            error_message = f"{source.name}: {stderr}"
            return False, original_size, 0, error_message

        converted_size = 0
        try:
            if target.exists():
                converted_size = target.stat().st_size
        except OSError:
            converted_size = 0

        return True, original_size, converted_size, None

    def _build_cwebp_command(self, source: Path, target: Path) -> List[str]:
        cmd = [ENCODER_NAME, "-q",
               str(self.quality), str(source), "-o", str(target)]

        if source.suffix.lower() == ".png" and self.quality == 100:
            cmd.insert(1, "-lossless")

        return cmd

    def _replace_originals(self, batch: FolderBatch, outdir: Path) -> List[str]:
        self.status.emit(f"Replacing originals in '{batch.folder.name}'...")
        errors: List[str] = []

        for original in batch.convertible_images:
            if not original.exists():
                continue
            try:
                original.unlink()
            except Exception as exc:
                message = f"Error deleting {original.name}: {exc}"
                self.status.emit(message)
                errors.append(message)

        for optimized in outdir.iterdir():
            destination = batch.folder / optimized.name
            try:
                shutil.move(str(optimized), destination)
            except Exception as exc:
                message = f"Error moving {optimized.name}: {exc}"
                self.status.emit(message)
                errors.append(message)

        self._cleanup_dir(outdir)
        return errors

    def _create_archive(self, batch: FolderBatch, outdir: Path) -> tuple[Optional[Path], Optional[int], List[str]]:
        archive_ext = ".cbz" if self.use_cbz else ".zip"
        archive_label = "CBZ" if self.use_cbz else "ZIP"
        archive_name = f"{batch.folder.name}{archive_ext}"
        archive_path = batch.folder.parent / archive_name
        errors: List[str] = []

        self.status.emit(
            f"Creating {archive_label} archive for '{batch.folder.name}'...")

        if archive_path.exists():
            try:
                archive_path.unlink()
            except Exception as exc:
                message = f"Unable to remove existing archive {archive_name}: {exc}"
                self.status.emit(message)
                errors.append(message)
                self._cleanup_dir(outdir)
                return None, None, errors

        try:
            optimized_files = list(outdir.iterdir())
            with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
                for optimized in optimized_files:
                    zipf.write(optimized, arcname=optimized.name)

                if self.skip_webp:
                    for original in batch.skipped_webp:
                        zipf.write(original, arcname=original.name)
        except Exception as exc:
            message = f"Error creating archive for '{batch.folder.name}': {exc}"
            self.status.emit(message)
            errors.append(message)
            self._cleanup_dir(outdir)
            return None, None, errors

        archive_size: Optional[int] = None
        try:
            if archive_path.exists():
                archive_size = archive_path.stat().st_size
        except OSError as exc:
            message = f"Unable to read archive size for '{archive_name}': {exc}"
            self.status.emit(message)
            errors.append(message)

        self._cleanup_dir(outdir)
        return archive_path if archive_path.exists() else None, archive_size, errors

    def _cleanup_dir(self, directory: Path) -> None:
        if not directory.exists():
            return
        try:
            shutil.rmtree(directory)
        except Exception as exc:
            self.status.emit(f"Error removing folder {directory}: {exc}")

    def _emit_progress(self, processed: int, total: int) -> None:
        if total <= 0:
            return

        percentage = min(100, int((processed / total) * 100))
        self.progress.emit(percentage)

    def _build_run_summary(
        self,
        folder_summaries: List[FolderSummary],
        run_start: float,
        total_work_units: int,
        processed_units: int,
        expected_conversions: int,
        cancelled: bool = False,
    ) -> Dict[str, Any]:
        duration = max(0.0, time.perf_counter() - run_start)
        total_converted = sum(
            summary.converted for summary in folder_summaries)
        total_skipped = sum(
            summary.skipped_existing for summary in folder_summaries)
        total_errors = sum(len(summary.errors) for summary in folder_summaries)
        total_bytes_original = sum(
            summary.bytes_original for summary in folder_summaries)
        total_bytes_converted = sum(
            summary.bytes_converted for summary in folder_summaries)
        total_bytes_saved = max(
            0, total_bytes_original - total_bytes_converted)
        archives_created = sum(
            1 for summary in folder_summaries if summary.archive_path)

        return {
            "cancelled": cancelled,
            "duration_seconds": duration,
            "total_images": total_work_units,
            "processed_images": processed_units,
            "expected_conversions": expected_conversions,
            "totals": {
                "converted": total_converted,
                "skipped_existing": total_skipped,
                "errors": total_errors,
                "bytes_original": total_bytes_original,
                "bytes_converted": total_bytes_converted,
                "bytes_saved": total_bytes_saved,
                "archives": archives_created,
            },
            "folders": [asdict(summary) for summary in folder_summaries],
        }


# Fix taskbar icon for Windows when running .pyw
if sys.platform == "win32":
    myappid = "ceavan.image_optimizer.1.0"  # arbitrary unique ID
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Image to WebP Optimizer")
        self.setWindowIcon(QIcon("icon.png"))
        self.resize(640, 420)
        self.cancel_requested = False

        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)
        self.setLayout(layout)

        header_label = QLabel("Image to WebP Optimizer")
        header_label.setObjectName("header-label")
        header_font = QFont()
        header_font.setPointSize(16)
        header_font.setBold(True)
        header_label.setFont(header_font)
        layout.addWidget(header_label)

        intro_label = QLabel(
            "Convert/compress large image folders into WebP files, with archive and replacement options."
        )
        intro_label.setObjectName("intro-label")
        intro_label.setWordWrap(True)
        layout.addWidget(intro_label)

        content_layout = QHBoxLayout()
        content_layout.setSpacing(16)
        layout.addLayout(content_layout)

        left_column_container = QWidget()
        left_column_container.setMaximumWidth(460)
        left_column_container.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding
        )
        left_column = QVBoxLayout(left_column_container)
        left_column.setSpacing(12)
        content_layout.addWidget(left_column_container)

        source_group = QGroupBox("Source")
        source_layout = QVBoxLayout()
        source_layout.setSpacing(8)
        source_group.setLayout(source_layout)
        left_column.addWidget(source_group)

        self.folder_display = QLabel("No folder selected yet.")
        self.folder_display.setObjectName("folder-display")
        self.folder_display.setWordWrap(True)
        source_layout.addWidget(self.folder_display)

        source_controls = QHBoxLayout()
        source_controls.setSpacing(8)
        self.btn_select = QPushButton("Browse…")
        self.btn_select.clicked.connect(self.select_folder)
        source_controls.addWidget(self.btn_select)
        source_controls.addStretch()
        source_layout.addLayout(source_controls)

        self.cb_library = QCheckBox(
            "Library Mode (Process each subfolder individually)")
        self.cb_library.setToolTip(
            "Optimize every subfolder instead of the selected folder.")
        source_layout.addWidget(self.cb_library)

        settings_group = QGroupBox("Conversion Settings")
        settings_layout = QVBoxLayout()
        settings_layout.setSpacing(8)
        settings_group.setLayout(settings_layout)
        left_column.addWidget(settings_group)

        quality_layout = QHBoxLayout()
        quality_layout.setSpacing(10)
        quality_label = QLabel("Quality")
        quality_layout.addWidget(quality_label)

        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(10, 100)
        self.quality_slider.setValue(75)
        self.quality_slider.setTickInterval(5)
        self.quality_slider.setSingleStep(1)
        quality_layout.addWidget(self.quality_slider, 1)

        self.quality_spin = QSpinBox()
        self.quality_spin.setRange(10, 100)
        self.quality_spin.setValue(self.quality_slider.value())
        quality_layout.addWidget(self.quality_spin)

        self.quality_slider.valueChanged.connect(self.quality_spin.setValue)
        self.quality_spin.valueChanged.connect(self.quality_slider.setValue)

        settings_layout.addLayout(quality_layout)

        quality_hint = QLabel(
            "Higher quality increases file size. Recommended: 70–80 for most artwork.")
        quality_hint.setStyleSheet("color: #888888; font-size: 10pt;")
        quality_hint.setWordWrap(True)
        settings_layout.addWidget(quality_hint)

        output_group = QGroupBox("Output Options")
        output_layout = QVBoxLayout()
        output_layout.setSpacing(6)
        output_group.setLayout(output_layout)
        left_column.addWidget(output_group)

        self.cb_replace = QCheckBox(
            "Replace originals with optimized WebP files (no archive)")
        self.cb_replace.setToolTip(
            "Deletes originals after conversion and moves WebP files into place.")
        output_layout.addWidget(self.cb_replace)

        self.cb_cbz = QCheckBox(
            "Create .cbz archive (comic book archive format)")
        self.cb_cbz.setToolTip(
            "Use when you want a CBZ instead of a standard ZIP archive.")
        output_layout.addWidget(self.cb_cbz)

        self.cb_skip_webp = QCheckBox("Skip images that are already WebP")
        self.cb_skip_webp.setChecked(True)
        output_layout.addWidget(self.cb_skip_webp)

        self.cb_replace.toggled.connect(self._toggle_archive_options)

        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(8)
        progress_group.setLayout(progress_layout)
        left_column.addWidget(progress_group)

        self.status_label = QLabel("Idle")
        self.status_label.setObjectName("status-label")
        self.status_label.setWordWrap(True)
        progress_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        progress_layout.addWidget(self.progress_bar)

        action_layout = QHBoxLayout()
        action_layout.addStretch()

        self.btn_start = QPushButton("Start Conversion")
        self.btn_start.clicked.connect(self.start_conversion)
        self.btn_start.setEnabled(False)
        action_layout.addWidget(self.btn_start)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.cancel_conversion)
        self.btn_cancel.setEnabled(False)
        action_layout.addWidget(self.btn_cancel)

        left_column.addLayout(action_layout)

        history_group = QGroupBox("Recent Runs")
        history_group.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        history_layout = QVBoxLayout()
        history_layout.setSpacing(8)
        history_group.setLayout(history_layout)
        content_layout.addWidget(history_group, 1)

        self.history_list = QListWidget()
        self.history_list.itemActivated.connect(
            self._on_history_item_activated)
        history_layout.addWidget(self.history_list, 1)

        history_buttons = QHBoxLayout()
        history_buttons.addStretch()
        self.btn_clear_history = QPushButton("Clear History")
        self.btn_clear_history.clicked.connect(self._on_clear_history_clicked)
        history_buttons.addWidget(self.btn_clear_history)
        history_layout.addLayout(history_buttons)

        self.selected_folder = None
        self.worker = None
        self.history_entries: List[Dict[str, Any]] = self._load_history()
        self.latest_summary: Optional[Dict[str, Any]] = None
        self.active_run_context: Optional[Dict[str, Any]] = None
        self._refresh_history_view()

        self._update_stylesheet()
        app = QApplication.instance()
        if app is not None:
            try:
                # type: ignore[attr-defined]
                app.paletteChanged.connect(self._update_stylesheet)
            except AttributeError:
                pass

        self._set_controls_enabled(is_running=False)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.selected_folder = Path(folder)
            self.folder_display.setText(f"Selected folder:\n{folder}")
            self.status_label.setText("Ready to start")
            self._set_controls_enabled(is_running=False)

    def start_conversion(self):
        if not self.selected_folder:
            QMessageBox.warning(self, "No folder",
                                "Please select a folder first!")
            return

        quality = self.quality_spin.value()
        use_cbz = self.cb_cbz.isChecked()
        replace_originals = self.cb_replace.isChecked()
        library_mode = self.cb_library.isChecked()
        self.progress_bar.setValue(0)

        if library_mode:
            subfolders = [f for f in self.selected_folder.iterdir()
                          if f.is_dir()]
            if not subfolders:
                QMessageBox.warning(
                    self, "No subfolders", "Selected folder has no subfolders to process!")
                return
            folders_to_process = subfolders
        else:
            folders_to_process = [self.selected_folder]

        skip_webp = self.cb_skip_webp.isChecked()
        self.active_run_context = {
            "selected_folder": str(self.selected_folder),
            "folders": [str(path) for path in folders_to_process],
            "library_mode": library_mode,
            "quality": quality,
            "use_cbz": use_cbz,
            "replace_originals": replace_originals,
            "skip_webp": skip_webp,
        }

        self.cancel_requested = False
        self._set_controls_enabled(is_running=True)
        self.status_label.setText("Preparing files...")
        self.latest_summary = None

        self.worker = Worker(folders_to_process, quality,
                             use_cbz, replace_originals, skip_webp)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.summary_ready.connect(self._on_worker_summary)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def cancel_conversion(self):
        if not self.worker or not self.worker.isRunning():
            return
        self.cancel_requested = True
        self.status_label.setText("Attempting to cancel...")
        self.worker.requestInterruption()
        self.btn_cancel.setEnabled(False)

    def _toggle_archive_options(self, replace_checked: bool) -> None:
        self.cb_cbz.setEnabled(not replace_checked)
        if replace_checked and self.cb_cbz.isChecked():
            self.cb_cbz.setChecked(False)

    def _set_controls_enabled(self, is_running: bool) -> None:
        self.btn_start.setEnabled(
            not is_running and self.selected_folder is not None)
        self.btn_select.setEnabled(not is_running)
        self.cb_library.setEnabled(not is_running)
        self.cb_cbz.setEnabled(
            not is_running and not self.cb_replace.isChecked())
        self.cb_replace.setEnabled(not is_running)
        self.cb_skip_webp.setEnabled(not is_running)
        self.quality_spin.setEnabled(not is_running)
        self.quality_slider.setEnabled(not is_running)
        self.btn_cancel.setEnabled(is_running)

    def _on_worker_summary(self, summary: Dict[str, Any]) -> None:
        if isinstance(summary, dict):
            self.latest_summary = summary

    def _update_stylesheet(self, palette: Optional[QPalette] = None) -> None:
        palette = palette or self.palette()

        window_color = palette.color(QPalette.ColorRole.Window)
        text_color = palette.color(QPalette.ColorRole.WindowText)
        is_dark = window_color.lightness() < 128

        muted = QColor(text_color)
        if is_dark:
            muted = muted.lighter(150)
        else:
            muted = muted.darker(130)

        header_color = QColor(text_color)
        if is_dark:
            header_color = header_color.lighter(130)
        else:
            header_color = header_color.darker(110)

        border_role = QPalette.ColorRole.Midlight if is_dark else QPalette.ColorRole.Mid
        border_color = palette.color(border_role)

        stylesheet = f"""
            QWidget {{
                font-size: 11pt;
            }}
            QGroupBox {{
                border: 1px solid {border_color.name()};
                border-radius: 6px;
                margin-top: 12px;
                padding: 12px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
            #header-label {{
                color: {header_color.name()};
            }}
            #intro-label, #folder-display {{
                color: {muted.name()};
            }}
            #status-label {{
                font-weight: 600;
            }}
            QPushButton {{
                min-height: 34px;
            }}
        """

        self.setStyleSheet(stylesheet)

    def _on_history_item_activated(self, item: QListWidgetItem) -> None:
        entry = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(entry, dict):
            return

        folder_path = entry.get("primary_folder")
        if not folder_path:
            return

        folder = Path(folder_path)
        if not folder.exists():
            QMessageBox.warning(
                self,
                "Missing folder",
                f"The folder '{folder_path}' could not be found on disk."
            )
            return

        self.selected_folder = folder
        self.folder_display.setText(f"Selected folder:\n{folder_path}")
        self.cb_library.setChecked(bool(entry.get("library_mode", False)))
        self.cb_cbz.setChecked(bool(entry.get("use_cbz", False)))
        self.cb_replace.setChecked(bool(entry.get("replace_originals", False)))
        self.cb_skip_webp.setChecked(bool(entry.get("skip_webp", True)))

        quality_value = entry.get("quality")
        if isinstance(quality_value, int):
            self.quality_spin.setValue(quality_value)

        self.status_label.setText("Ready to start (loaded from history)")
        self._set_controls_enabled(is_running=False)

    def _on_clear_history_clicked(self) -> None:
        if not self.history_entries:
            return

        confirm = QMessageBox.question(
            self,
            "Clear history",
            "Remove all saved runs from history?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.history_entries.clear()
            self._save_history()
            self._refresh_history_view()

    def _load_history(self) -> List[Dict[str, Any]]:
        if not HISTORY_FILE.exists():
            return []
        try:
            raw_data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except Exception:
            return []

        if not isinstance(raw_data, list):
            return []

        entries: List[Dict[str, Any]] = []
        for item in raw_data[:MAX_HISTORY_ENTRIES]:
            if isinstance(item, dict):
                entries.append(item)
        return entries

    def _save_history(self) -> None:
        try:
            HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
            HISTORY_FILE.write_text(
                json.dumps(
                    self.history_entries[:MAX_HISTORY_ENTRIES], indent=2),
                encoding="utf-8"
            )
        except Exception:
            # Saving history is non-critical; ignore errors silently.
            return

    def _refresh_history_view(self) -> None:
        self.history_list.clear()
        for entry in self.history_entries:
            label = self._format_history_entry(entry)
            list_item = QListWidgetItem(label)
            list_item.setData(Qt.ItemDataRole.UserRole, entry)
            self.history_list.addItem(list_item)

    def _format_history_entry(self, entry: Dict[str, Any]) -> str:
        timestamp = entry.get("timestamp", "")
        try:
            timestamp_dt = datetime.fromisoformat(timestamp)
            timestamp_display = timestamp_dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            timestamp_display = "Unknown time"

        folder_display = entry.get("primary_folder") or "Unknown folder"
        totals = entry.get("totals", {})
        converted = totals.get("converted", 0)
        bytes_saved = totals.get("bytes_saved")
        size_display = ""
        if isinstance(bytes_saved, int) and bytes_saved > 0:
            size_display = f" • {self._human_size(bytes_saved)} saved"

        return f"{timestamp_display} • {folder_display} • {converted} files{size_display}"

    def _human_size(self, num_bytes: int) -> str:
        units = ["B", "KB", "MB", "GB", "TB"]
        size = float(num_bytes)
        unit_index = 0
        while size >= 1024 and unit_index < len(units) - 1:
            size /= 1024
            unit_index += 1
        return f"{size:.1f} {units[unit_index]}"

    def _percent_saved(self, original_bytes: int, converted_bytes: int) -> str:
        if original_bytes <= 0:
            return "—"
        saved = max(0, original_bytes - converted_bytes)
        percent = (saved / original_bytes) * 100
        return f"{percent:.1f}%"

    def _append_history_entry(self, summary: Dict[str, Any]) -> None:
        if summary.get("cancelled"):
            return

        options = (self.active_run_context or {}).copy()
        folders = summary.get("folders") or []
        primary_folder = options.get("selected_folder")
        if not primary_folder and folders:
            primary_folder = folders[0].get("folder")

        entry = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "primary_folder": primary_folder,
            "folders": options.get("folders") or [f.get("folder") for f in folders],
            "totals": summary.get("totals", {}),
            "duration_seconds": summary.get("duration_seconds"),
            "processed_images": summary.get("processed_images"),
            "expected_conversions": summary.get("expected_conversions"),
            "quality": options.get("quality"),
            "use_cbz": options.get("use_cbz"),
            "replace_originals": options.get("replace_originals"),
            "library_mode": options.get("library_mode"),
            "skip_webp": options.get("skip_webp"),
        }

        # Ensure strings for JSON serialization
        if isinstance(entry["folders"], list):
            entry["folders"] = [
                str(path) if path is not None else "" for path in entry["folders"]]
        if entry["primary_folder"]:
            entry["primary_folder"] = str(entry["primary_folder"])

        self.history_entries.insert(0, entry)
        del self.history_entries[MAX_HISTORY_ENTRIES:]
        self._save_history()
        self._refresh_history_view()

    def _show_results_dialog(self, summary: Dict[str, Any]) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Conversion Results")

        layout = QVBoxLayout(dialog)
        totals = summary.get("totals", {})
        converted = totals.get("converted", 0)
        skipped = totals.get("skipped_existing", 0)
        archives = totals.get("archives", 0)
        bytes_saved = totals.get("bytes_saved", 0)
        duration = summary.get("duration_seconds", 0.0)

        headline = QLabel(
            f"Converted {converted} image(s) in {duration:.1f} seconds."
        )
        headline.setWordWrap(True)
        layout.addWidget(headline)

        totals_original = totals.get("bytes_original", 0)
        totals_converted = totals.get("bytes_converted", 0)
        subline_parts = []
        if skipped:
            subline_parts.append(f"Skipped {skipped} existing WebP file(s)")
        if totals_original and totals_converted:
            percent_saved = self._percent_saved(
                totals_original, totals_converted)
            subline_parts.append(
                f"Original: {self._human_size(totals_original)} → {self._human_size(totals_converted)} "
                f"({percent_saved} saved)"
            )
        elif bytes_saved:
            subline_parts.append(
                f"Estimated savings: {self._human_size(int(bytes_saved))}")
        if archives:
            subline_parts.append(f"Created {archives} archive(s)")

        if subline_parts:
            subline = QLabel("\n".join(subline_parts))
            subline.setStyleSheet("color: #666666;")
            subline.setWordWrap(True)
            layout.addWidget(subline)

        table = QTableWidget()
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels(
            ["Folder", "Converted", "Skipped", "Original Size",
                "Result Size", "% Saved", "Errors"]
        )
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        folder_rows = summary.get("folders") or []
        for row_index, folder_info in enumerate(folder_rows):
            table.insertRow(row_index)
            folder_path = folder_info.get("folder", "Unknown")
            folder_name = Path(folder_path).name or folder_path
            folder_item = QTableWidgetItem(folder_name)
            folder_item.setToolTip(folder_path)
            table.setItem(row_index, 0, folder_item)

            converted_value = folder_info.get("converted", 0)
            converted_item = QTableWidgetItem(str(converted_value))
            converted_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_index, 1, converted_item)

            skipped_value = folder_info.get("skipped_existing", 0)
            skipped_item = QTableWidgetItem(str(skipped_value))
            skipped_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_index, 2, skipped_item)

            original_bytes = max(0, int(folder_info.get("bytes_original", 0)))
            converted_bytes = max(
                0, int(folder_info.get("bytes_converted", 0)))
            saved_value = max(
                0,
                int(folder_info.get("bytes_original", 0)) -
                int(folder_info.get("bytes_converted", 0))
            )
            original_item = QTableWidgetItem(self._human_size(
                original_bytes) if original_bytes else "—")
            original_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_index, 3, original_item)

            converted_item = QTableWidgetItem(self._human_size(
                converted_bytes) if converted_bytes else "—")
            converted_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_index, 4, converted_item)

            percent_item = QTableWidgetItem(
                self._percent_saved(
                    original_bytes, converted_bytes) if original_bytes else "—"
            )
            percent_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row_index, 5, percent_item)

            errors = folder_info.get("errors") or []
            errors_item = QTableWidgetItem(str(len(errors)))
            errors_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if errors:
                errors_item.setToolTip("\n".join(errors[:5]))
            table.setItem(row_index, 6, errors_item)

        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col in range(1, 7):
            table.horizontalHeader().setSectionResizeMode(
                col, QHeaderView.ResizeMode.ResizeToContents)

        layout.addWidget(table)

        # Aggregate error details if present
        all_errors = [
            err for folder in folder_rows for err in folder.get("errors", [])]
        if all_errors:
            errors_label = QLabel(
                "Issues encountered:\n" + "\n".join(all_errors[:5]))
            errors_label.setWordWrap(True)
            errors_label.setStyleSheet("color: #b22222;")
            layout.addWidget(errors_label)
            if len(all_errors) > 5:
                more_label = QLabel(f"...and {len(all_errors) - 5} more.")
                more_label.setStyleSheet("color: #b22222;")
                layout.addWidget(more_label)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(dialog.reject)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.resize(720, 420)
        dialog.exec()

    def on_finished(self):
        summary = self.latest_summary or {}
        was_cancelled = self.cancel_requested or bool(summary.get("cancelled"))
        self._set_controls_enabled(is_running=False)
        self.worker = None

        if was_cancelled:
            self.status_label.setText("Conversion cancelled.")
        else:
            totals = summary.get("totals", {})
            converted_total = totals.get("converted", 0)
            if summary and (summary.get("folders") or summary.get("total_images", 0) > 0):
                self.status_label.setText(
                    f"Completed {converted_total} image(s).")
                self._append_history_entry(summary)
                self._show_results_dialog(summary)
            else:
                QMessageBox.information(
                    self, "Done", "No images were converted.")
                self.status_label.setText("Idle")
        self.cancel_requested = False
        self.latest_summary = None
        self.active_run_context = None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("icon.png"))
    window = App()
    window.show()
    sys.exit(app.exec())
