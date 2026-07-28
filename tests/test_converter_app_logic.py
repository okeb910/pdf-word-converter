import os
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pdf_word_converter as converter
from engine_models import EngineState, EngineStatus


class ValueHolder:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class ConverterAppLogicTests(unittest.TestCase):
    @staticmethod
    def _dense_table_matrix(shape, text):
        rows, columns = shape
        matrix = [[""] * columns for _row in range(rows)]
        if rows and columns:
            matrix[0][0] = text
        return tuple(tuple(row) for row in matrix)

    @staticmethod
    def _horizontal_table_spans(shape, matrix):
        rows, columns = shape
        spans = []
        for row_index in range(rows):
            starts = [
                column_index
                for column_index, value in enumerate(matrix[row_index])
                if value is not None
            ]
            for start_index, column_index in enumerate(starts):
                next_column = (
                    starts[start_index + 1]
                    if start_index + 1 < len(starts)
                    else columns
                )
                spans.append(
                    (
                        row_index,
                        column_index,
                        1,
                        next_column - column_index,
                    )
                )
        return tuple(spans)

    @staticmethod
    def _empty_table_border_edges(spans):
        return tuple((False, False, False, False) for _span in spans)

    @classmethod
    def _table_double(cls, **values):
        values.setdefault(
            "table_cell_matrices",
            tuple(
                cls._dense_table_matrix(shape, text)
                for text, shape in zip(
                    values.get("table_texts", ()) or (),
                    values.get("table_shapes", ()) or (),
                )
            ),
        )
        values.setdefault(
            "table_cell_spans",
            tuple(
                cls._horizontal_table_spans(shape, matrix)
                for shape, matrix in zip(
                    values.get("table_shapes", ()) or (),
                    values.get("table_cell_matrices", ()) or (),
                )
            ),
        )
        values.setdefault(
            "table_cell_border_edges",
            tuple(
                cls._empty_table_border_edges(spans)
                for spans in values.get("table_cell_spans", ()) or ()
            ),
        )
        values.setdefault(
            "table_column_widths",
            tuple(
                (1.0,) * int(shape[1])
                for shape in values.get("table_shapes", ()) or ()
            ),
        )
        return SimpleNamespace(**values)

    @classmethod
    def _fidelity_report(cls, path, **overrides):
        values = {
            "path": Path(path),
            "page_count": 1,
            "analyzed_pages": 1,
            "table_count": 0,
            "table_text_character_count": 0,
            "table_shapes": (),
            "table_cell_counts": (),
            "table_texts": (),
            "table_cell_matrices": (),
            "table_cell_spans": (),
            "table_cell_border_edges": (),
            "table_column_widths": (),
            "table_text_extraction_failure_count": 0,
            "selectable_text_character_count": 0,
            "editable_table_candidate": False,
            "table_layout_suspected": False,
            "tagged_table_structure_present": False,
            "widget_count": 0,
            "checkbox_symbol_count": 0,
            "vector_mark_count": 0,
            "symbol_font_run_count": 0,
            "vector_path_count": 0,
            "is_complex": False,
            "reasons": (),
        }
        values.update(overrides)
        if (
            "table_cell_matrices" not in overrides
            and values["table_texts"]
            and values["table_shapes"]
        ):
            values["table_cell_matrices"] = tuple(
                cls._dense_table_matrix(shape, text)
                for text, shape in zip(
                    values["table_texts"],
                    values["table_shapes"],
                )
            )
        if (
            "table_cell_spans" not in overrides
            and values["table_cell_matrices"]
            and values["table_shapes"]
        ):
            values["table_cell_spans"] = tuple(
                cls._horizontal_table_spans(shape, matrix)
                for shape, matrix in zip(
                    values["table_shapes"],
                    values["table_cell_matrices"],
                )
            )
        if (
            "table_cell_border_edges" not in overrides
            and values["table_cell_spans"]
        ):
            values["table_cell_border_edges"] = tuple(
                cls._empty_table_border_edges(spans)
                for spans in values["table_cell_spans"]
            )
        if "table_column_widths" not in overrides and values["table_shapes"]:
            values["table_column_widths"] = tuple(
                (1.0,) * int(shape[1])
                for shape in values["table_shapes"]
            )
        return converter.PdfFidelityRisk(**values)

    @staticmethod
    def _copy_test_snapshot(source, destination, _expected_identity):
        destination_path = Path(destination)
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        destination_path.write_bytes(f"snapshot:{Path(source).name}".encode("utf-8"))
        return destination_path

    @staticmethod
    def _bind_policy_snapshots(policy, *sources):
        for source in map(Path, sources):
            identity = converter._source_file_identity(source)
            snapshot_directory = tempfile.TemporaryDirectory(
                prefix="pdf-word-converter-test-preflight-"
            )
            snapshot_path = Path(snapshot_directory.name) / source.name
            converter._copy_verified_source_snapshot(
                source,
                snapshot_path,
                identity,
            )
            policy.source_identities[source] = identity
            policy.source_snapshots[source] = snapshot_path
            policy.snapshot_directories.append(snapshot_directory)
        return policy

    @staticmethod
    def _app_for_pdf_fidelity_choice(method):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.input_paths = [Path("anonymous-complex-report.pdf")]
        app.is_macos = False
        app.method_var = ValueHolder(method)
        app.log = Mock()
        app._method_availability = Mock(return_value=True)
        app.root = SimpleNamespace(
            after=lambda _delay, callback, *args: callback(*args)
        )
        app._cancel_event = threading.Event()
        app.progress_text_var = ValueHolder("")
        app._ask_yes_no_cancel_from_worker = Mock()
        app._ask_ok_cancel_from_worker = Mock()
        app._source_file_identity = Mock(
            return_value=(1, 2, 3, 4, "test-digest")
        )
        app._copy_verified_source_snapshot = Mock(
            side_effect=ConverterAppLogicTests._copy_test_snapshot
        )
        return app

    def test_cancelled_source_switch_does_not_change_pdf_target(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = "pdf"
        app.input_paths = [Path("current.pdf")]
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app._clear_queue = Mock()
        app._refresh_queue = Mock()
        app.log = Mock()

        with patch.object(
            converter.filedialog, "askopenfilenames", return_value=("slides.pptx",)
        ), patch.object(converter.messagebox, "askyesno", return_value=False):
            app.select_files("powerpoint")

        self.assertEqual(app._current_target_kind(), "word")
        self.assertEqual(app.source_kind, "pdf")
        app._clear_queue.assert_not_called()

    def test_powerpoint_selection_accepts_ppt_and_pptx(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = None
        app.input_paths = []
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app._refresh_queue = Mock()
        app.log = Mock()

        with patch.object(
            converter.filedialog,
            "askopenfilenames",
            return_value=("legacy.ppt", "modern.pptx", "ignored.pdf"),
        ):
            app.select_files("powerpoint")

        self.assertEqual(app.source_kind, "powerpoint")
        self.assertEqual(app._current_target_kind(), "pdf")
        self.assertEqual([path.suffix for path in app.input_paths], [".ppt", ".pptx"])

    def test_engine_status_update_never_prompts_for_office_on_startup(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._method_labels = {
            "word_com": Mock(),
            "images": Mock(),
            "libreoffice": Mock(),
        }
        app.pptx_status_label = Mock()
        app.method_var = ValueHolder("images")
        app.engine_status_var = ValueHolder("")
        app._update_method_panel = Mock()
        app._update_action_states = Mock()
        app.log = Mock()
        app._prompt_word_install = Mock()

        app._update_engine_status(False, False, False, True, True, None)

        app._prompt_word_install.assert_not_called()
        self.assertIn("PowerPoint: 不可用", app.engine_status_var.get())

    def test_engine_status_update_does_not_treat_failure_objects_as_available(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._method_labels = {
            converter.WORD_NATIVE: Mock(),
            "images": Mock(),
            "libreoffice": Mock(),
        }
        app.pptx_status_label = Mock()
        app.method_var = ValueHolder("images")
        app.engine_status_var = ValueHolder("")
        app._update_method_panel = Mock()
        app._update_action_states = Mock()
        app.log = Mock()

        app._update_engine_status(
            EngineStatus(EngineState.MISSING),
            EngineStatus(EngineState.LAUNCH_FAILED),
            EngineStatus(EngineState.TIMEOUT),
            EngineStatus(EngineState.AVAILABLE),
            EngineStatus(EngineState.AVAILABLE),
            None,
        )

        self.assertFalse(app._method_availability(converter.WORD_NATIVE))
        self.assertFalse(app._method_availability(converter.POWERPOINT_NATIVE))
        self.assertFalse(app._method_availability("libreoffice"))
        self.assertIn("PowerPoint: 已安装/启动失败", app.engine_status_var.get())

    def test_windows_converter_rejects_failed_structured_engine_statuses(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = False
        convert = app._create_converter(
            converter.resolve_conversion_spec("word", "pdf"),
            method=None,
        )

        with patch.object(
            converter,
            "word_com_available",
            return_value=EngineStatus(EngineState.LAUNCH_FAILED),
        ), patch.object(
            converter,
            "libreoffice_available",
            return_value=EngineStatus(EngineState.TIMEOUT),
        ), patch.object(converter, "docx_to_pdf_via_word") as word_convert, patch.object(
            converter, "docx_to_pdf_via_libreoffice"
        ) as libreoffice_convert, self.assertRaisesRegex(
            RuntimeError, "需要 Microsoft Word 或 LibreOffice"
        ):
            convert("source.docx", "output.pdf", Mock())

        word_convert.assert_not_called()
        libreoffice_convert.assert_not_called()

    def test_windows_native_failure_removes_partial_output_and_falls_back(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = False
        convert = app._create_converter(
            converter.resolve_conversion_spec("word", "pdf"),
            method=None,
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "result.pdf"

            def fail_native(_source, target, _progress):
                Path(target).write_bytes(b"partial-native")
                raise RuntimeError("Word COM failed")

            def succeed_fallback(_source, target, _progress):
                self.assertFalse(Path(target).exists())
                Path(target).write_bytes(b"libreoffice")

            with patch.object(
                converter,
                "word_com_available",
                return_value=EngineStatus(EngineState.UNVERIFIED),
            ), patch.object(
                converter,
                "libreoffice_available",
                return_value=EngineStatus(EngineState.UNVERIFIED),
            ), patch.object(
                converter, "docx_to_pdf_via_word", side_effect=fail_native
            ), patch.object(
                converter,
                "docx_to_pdf_via_libreoffice",
                side_effect=succeed_fallback,
            ):
                progress = Mock()
                convert("source.docx", output, progress)

            self.assertEqual(output.read_bytes(), b"libreoffice")
            self.assertTrue(
                any(
                    "改用 LibreOffice" in call.args[0] and call.args[1] == 5
                    for call in progress.call_args_list
                )
            )

    def test_pdf_to_word_editable_converters_never_run_legacy_postprocessing(self):
        cases = (
            (converter.WORD_NATIVE, "pdf_to_word_via_word"),
            ("libreoffice", "pdf_to_word_via_libreoffice"),
        )
        for method, converter_name in cases:
            with self.subTest(method=method):
                app = converter.ConverterApp.__new__(converter.ConverterApp)
                app.is_macos = False
                convert = app._create_converter(
                    converter.resolve_conversion_spec("pdf", "word"),
                    method=method,
                )

                with patch.object(converter, converter_name) as editable_convert, patch.object(
                    converter, "fix_converted_docx"
                ) as legacy_fix:
                    convert("anonymous.pdf", "anonymous.docx", Mock())

                editable_convert.assert_called_once()
                legacy_fix.assert_not_called()

    def test_legacy_docx_fix_preserves_bytes_and_missing_path_still_fails(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            output = folder / "anonymous.docx"
            original_bytes = b"anonymous engine output\x00\xff"
            output.write_bytes(original_bytes)
            progress = Mock()

            converter.fix_converted_docx(output, progress=progress)

            self.assertEqual(output.read_bytes(), original_bytes)
            progress.assert_called_once_with("已保留转换引擎原始排版", 95)

            missing = folder / "missing.docx"
            with self.assertRaises(FileNotFoundError):
                converter.fix_converted_docx(missing, progress=progress)
            self.assertEqual(progress.call_count, 1)

    def test_non_table_complex_pdf_choice_is_scoped_to_that_file(self):
        original_method = converter.WORD_NATIVE
        source = Path("anonymous-complex-report.pdf")
        report = self._fidelity_report(
            source,
            is_complex=True,
            reasons=("检测到表单图形", "检测到矢量勾选"),
        )
        cases = (
            (True, "images"),
            (False, None),
            (None, None),
        )

        for choice, expected_override in cases:
            with self.subTest(choice=choice):
                app = self._app_for_pdf_fidelity_choice(original_method)
                app._ask_yes_no_cancel_from_worker.return_value = choice
                with patch.object(
                    converter, "analyze_pdf_fidelity_risk", return_value=report
                ) as analyze:
                    policy = app._choose_pdf_word_fidelity_mode(original_method)

                analyze.assert_called_once()
                analyzed_snapshot = Path(analyze.call_args.args[0])
                self.assertIsNone(analyze.call_args.kwargs["max_pages"])
                cancel_requested = analyze.call_args.kwargs["cancel_requested"]
                self.assertTrue(callable(cancel_requested))
                self.assertFalse(cancel_requested())
                self.assertEqual(analyzed_snapshot.name, source.name)
                self.assertEqual(analyzed_snapshot.suffix, ".pdf")
                self.assertNotEqual(
                    app._path_key(analyzed_snapshot),
                    app._path_key(source),
                )
                app._ask_yes_no_cancel_from_worker.assert_called_once()
                self.assertEqual(app.method_var.get(), original_method)
                if choice is None:
                    self.assertIsNone(policy)
                    self.assertFalse(analyzed_snapshot.exists())
                else:
                    self.assertIsInstance(policy, converter.PdfWordBatchPolicy)
                    self.assertEqual(policy.default_method, original_method)
                    self.assertEqual(
                        policy.method_overrides.get(source),
                        expected_override,
                    )
                    self.assertEqual(set(policy.source_identities), {source})
                    self.assertEqual(
                        policy.source_snapshots[source],
                        analyzed_snapshot,
                    )
                    self.assertTrue(analyzed_snapshot.is_file())
                    policy.cleanup_snapshots()

    def test_pdf_fidelity_analysis_failure_is_blocked_fail_closed(self):
        original_method = "libreoffice"
        app = self._app_for_pdf_fidelity_choice(original_method)
        source = app.input_paths[0]

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            side_effect=RuntimeError("anonymous parser failure"),
        ):
            policy = app._choose_pdf_word_fidelity_mode(original_method)

        self.assertEqual(policy.default_method, original_method)
        self.assertEqual(policy.unverified_paths, {source})
        self.assertIn(source, policy.blocked_paths)
        self.assertEqual(policy.image_protected_paths, {source})
        app._ask_yes_no_cancel_from_worker.assert_not_called()
        app._ask_ok_cancel_from_worker.assert_called_once()
        self.assertTrue(
            any(
                "版式预检跳过" in call.args[0]
                and "anonymous parser failure" in call.args[0]
                for call in app.log.call_args_list
            )
        )
        self.assertTrue(
            any(
                "均无法检查" in call.args[0]
                and "标记失败" in call.args[0]
                for call in app.log.call_args_list
            )
        )
        self.assertFalse(
            any("未发现高风险结构" in call.args[0] for call in app.log.call_args_list)
        )



    def test_preflight_analysis_failure_reaches_batch_as_failed_without_output(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "unverified.pdf"
            source.write_bytes(b"not-a-real-pdf")
            app = self._app_for_pdf_fidelity_choice("libreoffice")
            app.input_paths = [source]
            app._ask_ok_cancel_from_worker.return_value = True
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app._create_converter = Mock()
            spec = converter.resolve_conversion_spec("pdf", "word")

            with patch.object(
                converter,
                "analyze_pdf_fidelity_risk",
                side_effect=RuntimeError("anonymous parser failure"),
            ):
                app._prepare_and_run_batch(
                    [source],
                    spec,
                    None,
                    "libreoffice",
                )

            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertFalse((source.parent / "unverified.docx").exists())
            app._create_converter.assert_not_called()
            app._on_batch_error.assert_not_called()

    def test_selectable_text_table_keeps_editable_method(self):
        original_method = converter.WORD_NATIVE
        app = self._app_for_pdf_fidelity_choice(original_method)
        source = Path("anonymous-editable-table.pdf")
        app.input_paths = [source]
        report = self._fidelity_report(
            source,
            is_complex=True,
            editable_table_candidate=True,
            table_count=1,
            table_text_character_count=128,
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            table_texts=("editable-table",),
            selectable_text_character_count=128,
            reasons=("检测到表格结构", "矢量路径较多"),
        )

        with patch.object(
            converter, "analyze_pdf_fidelity_risk", return_value=report
        ):
            policy = app._choose_pdf_word_fidelity_mode(original_method)

        self.assertEqual(policy.default_method, original_method)
        self.assertEqual(policy.method_overrides[source], original_method)
        self.assertEqual(policy.editable_table_reports[source], report)
        self.assertEqual(policy.blocked_paths, {})
        self.assertEqual(app.method_var.get(), original_method)
        app._ask_yes_no_cancel_from_worker.assert_not_called()
        app._ask_ok_cancel_from_worker.assert_not_called()
        self.assertTrue(
            any("已锁定可编辑转换" in call.args[0] for call in app.log.call_args_list)
        )

    def test_manual_image_mode_routes_editable_table_to_word_only(self):
        source = Path("anonymous-editable-table.pdf")
        app = self._app_for_pdf_fidelity_choice("images")
        app.input_paths = [source]
        app._method_availability.side_effect = (
            lambda method: method in {"images", converter.WORD_NATIVE}
        )
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            is_complex=True,
            editable_table_candidate=True,
            table_count=1,
            table_text_character_count=64,
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            table_texts=("editable-table",),
            selectable_text_character_count=64,
            reasons=("检测到表格结构",),
        )

        with patch.object(
            converter, "analyze_pdf_fidelity_risk", return_value=report
        ):
            policy = app._choose_pdf_word_fidelity_mode("images")

        self.assertEqual(policy.default_method, "images")
        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertEqual(policy.editable_table_reports[source], report)
        app._ask_yes_no_cancel_from_worker.assert_not_called()
        app._ask_ok_cancel_from_worker.assert_called_once()
        self.assertIn(
            "不能使用图片模式",
            app._ask_ok_cancel_from_worker.call_args.args[0],
        )

    def test_manual_image_mode_routes_any_selectable_text_to_editable_engine(self):
        source = Path("anonymous-selectable-text.pdf")
        app = self._app_for_pdf_fidelity_choice("images")
        app.input_paths = [source]
        app._method_availability.side_effect = (
            lambda method: method in {"images", converter.WORD_NATIVE}
        )
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            selectable_text_character_count=90,
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode("images")

        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertEqual(policy.image_protected_paths, {source})
        self.assertNotIn(source, policy.editable_table_reports)
        self.assertEqual(policy.selectable_text_reports[source], report)
        self.assertIn(
            "可选择文字不能使用图片模式",
            app._ask_ok_cancel_from_worker.call_args.args[0],
        )

    def test_unverifiable_text_mapping_converts_with_completion_warning(self):
        source = Path("unverifiable-font-map.pdf")
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [source]
        report = self._fidelity_report(
            source,
            selectable_text_character_count=1,
            selectable_text="\ufffd",
            unverifiable_text_character_count=1,
            is_complex=True,
            reasons=("检测到无法可靠映射的文字字符",),
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode(converter.WORD_NATIVE)

        self.assertNotIn(source, policy.blocked_paths)
        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertEqual(policy.selectable_text_reports[source], report)
        self.assertIn(source, policy.completion_warnings)
        self.assertIn("人工核对", policy.completion_warnings[source])

    def test_unconfirmed_table_layout_continues_with_completion_warning(self):
        source = Path("anonymous-unconfirmed-table.pdf")
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [source]
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            is_complex=True,
            table_layout_suspected=True,
            selectable_text_character_count=90,
            reasons=("版面呈现重复行列或网格，但表格结构无法确认",),
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode(converter.WORD_NATIVE)

        self.assertNotIn(source, policy.blocked_paths)
        self.assertEqual(policy.image_protected_paths, {source})
        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertNotIn(source, policy.editable_table_reports)
        self.assertEqual(policy.selectable_text_reports[source], report)
        self.assertIn(source, policy.advisory_validation_paths)
        self.assertIn("已优先继续转换", policy.completion_warnings[source])
        app._ask_ok_cancel_from_worker.assert_not_called()
        self.assertTrue(
            any(
                "疑似表格但无法提取" in call.args[0]
                for call in app.log.call_args_list
            )
        )

    def test_recognized_table_with_unconfirmed_region_uses_advisory_validation(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "recognized-and-unconfirmed-table.pdf"
            source.write_bytes(b"table-source")
            app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
            app.input_paths = [source]
            app._ask_ok_cancel_from_worker.return_value = True
            report = self._fidelity_report(
                source,
                is_complex=True,
                editable_table_candidate=True,
                table_count=1,
                table_text_character_count=16,
                table_shapes=((2, 2),),
                table_cell_counts=(4,),
                table_texts=("editable-table",),
                selectable_text_character_count=16,
                unconfirmed_table_region_count=1,
                reasons=(
                    "detected table structure",
                    "one additional table-like region is unconfirmed",
                ),
            )

            with patch.object(
                converter,
                "analyze_pdf_fidelity_risk",
                return_value=report,
            ):
                policy = app._choose_pdf_word_fidelity_mode(
                    converter.WORD_NATIVE
                )

            self.assertNotIn(source, policy.blocked_paths)
            self.assertEqual(policy.editable_table_reports[source], report)
            self.assertEqual(
                policy.method_overrides[source],
                converter.WORD_NATIVE,
            )
            self.assertIn(source, policy.advisory_validation_paths)
            self.assertIn("仔细检查", policy.completion_warnings[source])

    def test_table_extraction_uncertainty_still_requires_editable_table_output(self):
        source = Path("anonymous-uncertain-table.pdf")
        app = self._app_for_pdf_fidelity_choice("images")
        app.input_paths = [source]
        app._method_availability.side_effect = (
            lambda method: method in {"images", converter.WORD_NATIVE}
        )
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            is_complex=True,
            table_count=1,
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
            table_texts=("",),
            table_text_extraction_failure_count=1,
            reasons=("检测到表格结构", "表格单元格提取不完整"),
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode("images")

        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertEqual(policy.image_protected_paths, {source})
        self.assertEqual(policy.editable_table_reports[source], report)
        self.assertIn(source, policy.advisory_validation_paths)
        self.assertIn("逐格校验基线", policy.completion_warnings[source])
        app._ask_ok_cancel_from_worker.assert_called_once()

    def test_editable_table_with_unverifiable_form_elements_continues_with_warning(self):
        cases = (
            (
                {"widget_count": 1},
                ("检测到 PDF 表单控件",),
            ),
            (
                {"symbol_font_run_count": 1},
                ("检测到符号字体文本",),
            ),
            (
                {"vector_mark_count": 1},
                ("检测到疑似矢量复选或勾选标记",),
            ),
        )
        for overrides, reasons in cases:
            with self.subTest(overrides=overrides):
                source = Path("anonymous-editable-form-table.pdf")
                app = self._app_for_pdf_fidelity_choice(
                    converter.WORD_NATIVE
                )
                app.input_paths = [source]
                app._ask_ok_cancel_from_worker.return_value = True
                report = self._fidelity_report(
                    source,
                    is_complex=True,
                    editable_table_candidate=True,
                    table_count=1,
                    table_text_character_count=36,
                    table_shapes=((2, 2),),
                    table_cell_counts=(4,),
                    table_texts=("editable-form-table",),
                    selectable_text_character_count=36,
                    reasons=reasons,
                    **overrides,
                )

                with patch.object(
                    converter,
                    "analyze_pdf_fidelity_risk",
                    return_value=report,
                ):
                    policy = app._choose_pdf_word_fidelity_mode(
                        converter.WORD_NATIVE
                    )

                self.assertNotIn(source, policy.blocked_paths)
                self.assertIn(
                    "完成后请重点核对",
                    policy.completion_warnings[source],
                )
                self.assertEqual(
                    policy.method_overrides[source],
                    converter.WORD_NATIVE,
                )
                self.assertEqual(policy.editable_table_reports[source], report)
                self.assertIn(source, policy.advisory_validation_paths)
                self.assertEqual(
                    policy.image_protected_paths,
                    {source},
                )
                app._ask_yes_no_cancel_from_worker.assert_not_called()
                app._ask_ok_cancel_from_worker.assert_not_called()

    def test_unicode_checkmark_in_table_uses_strict_editable_gate(self):
        source = Path("anonymous-unicode-checkmark-table.pdf")
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [source]
        report = self._fidelity_report(
            source,
            is_complex=True,
            editable_table_candidate=True,
            table_count=1,
            table_text_character_count=3,
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
            table_texts=("☑同意",),
            table_cell_matrices=((("☑", "同意"),),),
            selectable_text_character_count=3,
            checkbox_symbol_count=1,
            reasons=("检测到表格结构", "检测到显式复选或勾选符号"),
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode(
                converter.WORD_NATIVE
            )

        self.assertNotIn(source, policy.blocked_paths)
        self.assertEqual(
            policy.method_overrides[source],
            converter.WORD_NATIVE,
        )
        self.assertEqual(policy.editable_table_reports[source], report)
        self.assertIn(source, policy.advisory_validation_paths)
        self.assertIn("勾选状态", policy.completion_warnings[source])
        app._ask_ok_cancel_from_worker.assert_not_called()
    def test_editable_pdf_waits_for_image_probe_before_start_is_enabled(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = "pdf"
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app.method_var = ValueHolder(converter.WORD_NATIVE)
        app._avail_methods = {
            converter.WORD_NATIVE: True,
            "images": None,
        }

        self.assertFalse(app._conversion_detection_ready())
        app._avail_methods["images"] = True
        self.assertTrue(app._conversion_detection_ready())

    def test_start_conversion_schedules_preflight_in_worker(self):
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.source_kind = "pdf"
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app.is_macos = False
        app._validate_output_dir = Mock(return_value=None)
        app.output_paths = []
        app._reset_queue_status = Mock()
        app._set_busy = Mock()
        app.progress = {"value": 0}
        worker = Mock()

        with patch.object(converter.threading, "Thread", return_value=worker) as thread:
            app.start_conversion()

        self.assertIs(
            thread.call_args.kwargs["target"].__func__,
            converter.ConverterApp._prepare_and_run_batch,
        )
        worker.start.assert_called_once_with()

    def test_preflight_cancel_stops_before_any_converter_runs(self):
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [Path("one.pdf"), Path("two.pdf")]
        app._run_batch = Mock()
        spec = converter.resolve_conversion_spec("pdf", "word")
        def cancel_after_first(snapshot, **_kwargs):
            app._cancel_event.set()
            return self._fidelity_report(snapshot)

        with patch.object(
            converter, "analyze_pdf_fidelity_risk", side_effect=cancel_after_first
        ):
            app._prepare_and_run_batch(app.input_paths, spec, None, converter.WORD_NATIVE)

        app._run_batch.assert_called_once_with(
            app.input_paths,
            spec,
            None,
            converter.WORD_NATIVE,
            None,
        )

    def test_preflight_cancel_exception_abandons_without_blocking_item(self):
        source = Path("cancel-during-analysis.pdf")
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [source]

        def cancel_during_analysis(
            _snapshot,
            *,
            max_pages,
            cancel_requested,
        ):
            self.assertIsNone(max_pages)
            self.assertFalse(cancel_requested())
            app._cancel_event.set()
            self.assertTrue(cancel_requested())
            raise converter.PdfFidelityAnalysisCancelled("cancelled")

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            side_effect=cancel_during_analysis,
        ):
            policy = app._choose_pdf_word_fidelity_mode(
                converter.WORD_NATIVE
            )

        self.assertIsNone(policy)
        log_messages = [call.args[0] for call in app.log.call_args_list]
        self.assertIn("版式预检已取消", log_messages)
        self.assertFalse(
            any("版式预检跳过" in message for message in log_messages)
        )

    def test_cancel_wait_timer_updates_each_second_and_stops_on_completion(self):
        callbacks = []
        cancelled_after_ids = []

        def after(delay, callback, *args):
            callbacks.append((delay, callback, args))
            return f"timer-{len(callbacks)}"

        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._is_converting = True
        app._cancel_event = threading.Event()
        app._close_after_batch = False
        app.cancel_btn = Mock()
        app.progress_text_var = ValueHolder("")
        app.queue_summary_var = ValueHolder("")
        app.progress = {"value": 0}
        app.output_paths = []
        app.log = Mock()
        app.root = SimpleNamespace(
            after=after,
            after_cancel=cancelled_after_ids.append,
        )
        app._set_busy = Mock(
            side_effect=lambda busy: setattr(app, "_is_converting", busy)
        )

        with patch.object(
            converter.time,
            "monotonic",
            side_effect=(100.0, 101.2),
        ), patch.object(converter.messagebox, "showinfo"):
            app.cancel_conversion()
            self.assertIn("已等待 0 秒", app.progress_text_var.get())
            self.assertTrue(app._cancel_event.is_set())
            app.cancel_btn.config.assert_called_once_with(state="disabled")
            self.assertEqual(callbacks[0][0], 1000)

            _delay, tick, args = callbacks.pop(0)
            tick(*args)
            self.assertIn("已等待 1 秒", app.progress_text_var.get())
            self.assertEqual(callbacks[0][0], 1000)
            active_after_id = app._cancel_wait_after_id

            app._on_batch_complete([])

        self.assertIn(active_after_id, cancelled_after_ids)
        self.assertIsNone(app._cancel_wait_after_id)
        self.assertIsNone(app._cancel_wait_started_at)
        self.assertEqual(app.progress_text_var.get(), "成功 0，失败 0，取消 0")

    def test_completion_warning_uses_warning_popup_after_success(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._close_after_batch = False
        app._stop_cancel_wait_timer = Mock()
        app._set_busy = Mock()
        app.output_paths = []
        app.queue_summary_var = ValueHolder("")
        app.progress_text_var = ValueHolder("")
        app.progress = {"value": 0}
        app.log = Mock()
        result = converter.BatchResult(
            Path("source.pdf"),
            Path("output.docx"),
            "success",
        )

        with patch.object(converter.messagebox, "showwarning") as showwarning, patch.object(
            converter.messagebox,
            "showinfo",
        ) as showinfo:
            app._on_batch_complete(
                [result],
                ["source.pdf：检测到 1 个字体映射可疑字符，请人工核对"],
            )

        showwarning.assert_called_once()
        showinfo.assert_not_called()
        self.assertEqual(showwarning.call_args.args[0], "转换完成，请仔细核对")
        self.assertIn("转换质量提醒", showwarning.call_args.args[1])
        self.assertIn("请人工核对", showwarning.call_args.args[1])

    def test_preflight_scans_every_file_and_all_pages(self):
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
        app.input_paths = [Path(f"anonymous-{index}.pdf") for index in range(12)]
        seen = []

        def analyze(snapshot, max_pages, cancel_requested):
            snapshot_path = Path(snapshot)
            seen.append(
                (
                    snapshot_path,
                    max_pages,
                    snapshot_path.is_file(),
                    snapshot_path.read_bytes(),
                    callable(cancel_requested),
                    cancel_requested(),
                )
            )
            return self._fidelity_report(snapshot_path)

        with patch.object(
            converter, "analyze_pdf_fidelity_risk", side_effect=analyze
        ):
            policy = app._choose_pdf_word_fidelity_mode(
                converter.WORD_NATIVE
            )

        self.assertIsInstance(policy, converter.PdfWordBatchPolicy)
        self.assertEqual(len(seen), len(app.input_paths))
        self.assertEqual(set(policy.source_identities), set(app.input_paths))
        self.assertEqual(set(policy.source_snapshots), set(app.input_paths))
        for source, (
            snapshot,
            max_pages,
            existed_during_analysis,
            content,
            has_cancel_callback,
            cancel_requested,
        ) in zip(
            app.input_paths,
            seen,
        ):
            self.assertIsNone(max_pages)
            self.assertTrue(has_cancel_callback)
            self.assertFalse(cancel_requested)
            self.assertTrue(existed_during_analysis)
            self.assertEqual(snapshot.name, source.name)
            self.assertEqual(snapshot.suffix, ".pdf")
            self.assertNotEqual(
                app._path_key(snapshot),
                app._path_key(source),
            )
            self.assertEqual(content, f"snapshot:{source.name}".encode("utf-8"))
            self.assertEqual(policy.source_snapshots[source], snapshot)
            self.assertTrue(snapshot.is_file())
        policy.cleanup_snapshots()
        self.assertTrue(all(not snapshot.exists() for snapshot, *_rest in seen))

    def test_manual_images_without_editable_engine_blocks_table_item(self):
        source = Path("anonymous-editable-table.pdf")
        app = self._app_for_pdf_fidelity_choice("images")
        app.input_paths = [source]
        app._method_availability.side_effect = lambda method: method == "images"
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            is_complex=True,
            editable_table_candidate=True,
            table_count=1,
            table_text_character_count=24,
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            table_texts=("editable-table",),
            selectable_text_character_count=24,
            reasons=("检测到表格结构",),
        )

        with patch.object(
            converter, "analyze_pdf_fidelity_risk", return_value=report
        ):
            policy = app._choose_pdf_word_fidelity_mode("images")

        self.assertIn(source, policy.blocked_paths)
        self.assertNotIn(source, policy.method_overrides)
        self.assertIn("不会生成照片", app._ask_ok_cancel_from_worker.call_args.args[1])

    def test_prompt_is_not_opened_when_cancel_arrives_before_main_callback(self):
        app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)

        def cancel_then_run(_delay, callback, *args):
            app._cancel_event.set()
            callback(*args)

        app.root = SimpleNamespace(after=cancel_then_run)
        del app._ask_yes_no_cancel_from_worker
        with patch.object(converter.messagebox, "askyesnocancel") as prompt:
            result = app._ask_yes_no_cancel_from_worker("title", "message")

        self.assertIsNone(result)
        prompt.assert_not_called()

    def test_mixed_batch_dispatches_table_and_graphic_to_different_engines(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            table_source = folder / "editable-table.pdf"
            graphic_source = folder / "graphic-form.pdf"
            table_source.write_bytes(b"table")
            graphic_source.write_bytes(b"graphic")
            report = self._table_double(
                table_count=1,
                table_text_character_count=20,
                table_texts=("X" * 20,),
                table_shapes=((1, 1),),
                table_cell_counts=(1,),
            )
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                method_overrides={
                    table_source: converter.WORD_NATIVE,
                    graphic_source: "images",
                },
                editable_table_reports={table_source: report},
            )
            self._bind_policy_snapshots(policy, table_source, graphic_source)
            expected_snapshots = dict(policy.source_snapshots)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )
            calls = []
            converted_sources = []

            def create(_spec, method, allow_image_fallback=True):
                calls.append((method, allow_image_fallback))

                def convert(_source, output, _progress):
                    converted_sources.append(Path(_source))
                    Path(output).write_bytes(method.encode("ascii"))

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_editable_table_output = Mock()
            spec = converter.resolve_conversion_spec("pdf", "word")

            with patch.object(
                converter,
                "repair_docx_table_topology",
                return_value=False,
            ) as repair:
                app._run_batch(
                    [table_source, graphic_source],
                    spec,
                    None,
                    converter.WORD_NATIVE,
                    policy,
                )

        self.assertEqual(
            calls,
            [
                (converter.WORD_NATIVE, False),
                ("images", True),
            ],
        )
        self.assertEqual(
            converted_sources,
            [
                expected_snapshots[table_source],
                expected_snapshots[graphic_source],
            ],
        )
        self.assertTrue(all(snapshot != source for snapshot, source in zip(
            converted_sources,
            (table_source, graphic_source),
        )))
        app._validate_editable_table_output.assert_called_once()
        repair.assert_called_once()
        results = app._on_batch_complete.call_args.args[0]
        self.assertEqual([result.status for result in results], ["success", "success"])
        app._on_batch_error.assert_not_called()

    def test_editable_table_repair_receives_source_border_evidence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "bordered-table.pdf"
            source.write_bytes(b"table")
            border_edges = (((True, True, True, True),),)
            report = self._table_double(
                table_count=1,
                table_text_character_count=5,
                table_texts=("Value",),
                table_shapes=((1, 1),),
                table_cell_counts=(1,),
                table_cell_border_edges=border_edges,
            )
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                editable_table_reports={source: report},
            )
            self._bind_policy_snapshots(policy, source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, _method, allow_image_fallback=True):
                self.assertFalse(allow_image_fallback)

                def convert(_source, output, _progress):
                    Path(output).write_bytes(b"editable")

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_editable_table_output = Mock()
            with patch.object(
                converter,
                "repair_docx_table_topology",
                return_value=False,
            ) as repair:
                app._run_batch(
                    [source],
                    converter.resolve_conversion_spec("pdf", "word"),
                    None,
                    converter.WORD_NATIVE,
                    policy,
                )

            repair.assert_called_once()
            self.assertEqual(
                repair.call_args.kwargs["table_cell_border_edges"],
                border_edges,
            )
            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(results[0].status, "success")
            app._on_batch_error.assert_not_called()

    def test_editable_table_progress_remains_monotonic_through_validation(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "table.pdf"
            source.write_bytes(b"table")
            report = self._table_double(
                table_count=1,
                table_text_character_count=5,
                table_texts=("Value",),
                table_shapes=((1, 1),),
                table_cell_counts=(1,),
            )
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                editable_table_reports={source: report},
            )
            self._bind_policy_snapshots(policy, source)

            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            progress_values = []
            app._progress_update = (
                lambda _current, _total, _message, pct: progress_values.append(pct)
            )
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, _method, allow_image_fallback=True):
                self.assertFalse(allow_image_fallback)

                def convert(_source, output, progress):
                    Path(output).write_bytes(b"editable")
                    progress("引擎完成", 100)

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_editable_table_output = Mock(
                side_effect=lambda _output, _report, progress: progress(
                    "表格校验完成",
                    99,
                )
            )
            with patch.object(
                converter,
                "repair_docx_table_topology",
                return_value=False,
            ) as repair:
                app._run_batch(
                    [source],
                    converter.resolve_conversion_spec("pdf", "word"),
                    None,
                    converter.WORD_NATIVE,
                    policy,
                )

            self.assertEqual(progress_values, sorted(progress_values))
            self.assertEqual(progress_values[-4:], [96, 97, 99, 100])
            repair.assert_called_once()
            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(results[0].status, "success")

    def test_advisory_table_validation_failure_keeps_output_and_warns(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "advisory-table.pdf"
            source.write_bytes(b"table")
            report = self._table_double(
                table_count=1,
                table_text_character_count=5,
                table_texts=("Value",),
                table_shapes=((1, 1),),
                table_cell_counts=(1,),
            )
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                editable_table_reports={source: report},
                advisory_validation_paths={source},
                completion_warnings={source: "源表格无法建立完整校验基线"},
            )
            self._bind_policy_snapshots(policy, source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, _method, allow_image_fallback=True):
                self.assertFalse(allow_image_fallback)

                def convert(_source, output, _progress):
                    Path(output).write_bytes(b"readable-docx-double")

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_editable_table_output = Mock(
                side_effect=RuntimeError("table structure mismatch")
            )
            with patch.object(
                converter,
                "inspect_editable_docx_tables",
                return_value=SimpleNamespace(),
            ), patch.object(
                converter,
                "repair_docx_table_topology",
                return_value=False,
            ):
                app._run_batch(
                    [source],
                    converter.resolve_conversion_spec("pdf", "word"),
                    None,
                    converter.WORD_NATIVE,
                    policy,
                )

            results, warnings = app._on_batch_complete.call_args.args
            self.assertEqual(results[0].status, "success")
            self.assertTrue(results[0].output.is_file())
            self.assertEqual(len(warnings), 1)
            self.assertIn("严格文字或结构校验未完全通过", warnings[0])
            app._on_batch_error.assert_not_called()

    def test_selectable_text_policy_disables_libreoffice_image_fallback(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "selectable-text.pdf"
            source.write_bytes(b"text")
            report = self._fidelity_report(
                source,
                selectable_text_character_count=4,
                selectable_text="Text",
            )
            policy = converter.PdfWordBatchPolicy(
                default_method="libreoffice",
                selectable_text_reports={source: report},
                image_protected_paths={source},
            )
            self._bind_policy_snapshots(policy, source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )
            calls = []

            def create(_spec, method, allow_image_fallback=True):
                calls.append((method, allow_image_fallback))

                def convert(_source, output, _progress):
                    Path(output).write_bytes(b"editable")

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_selectable_text_output = Mock()
            spec = converter.resolve_conversion_spec("pdf", "word")

            app._run_batch(
                [source],
                spec,
                None,
                "libreoffice",
                policy,
            )

        self.assertEqual(calls, [("libreoffice", False)])
        results = app._on_batch_complete.call_args.args[0]
        self.assertEqual(results[0].status, "success")
        app._validate_selectable_text_output.assert_called_once()
        app._on_batch_error.assert_not_called()

    def test_advisory_selectable_text_mismatch_keeps_output_and_warns(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "long-report.pdf"
            source.write_bytes(b"text")
            report = self._fidelity_report(
                source,
                page_count=225,
                analyzed_pages=225,
                selectable_text_character_count=100,
                selectable_text="X" * 100,
                table_analysis_limited=True,
            )
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                selectable_text_reports={source: report},
                image_protected_paths={source},
                advisory_validation_paths={source},
                completion_warnings={source: "长文档表格未逐页严格核对"},
            )
            self._bind_policy_snapshots(policy, source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, _method, allow_image_fallback=True):
                self.assertFalse(allow_image_fallback)

                def convert(_source, output, _progress):
                    Path(output).write_bytes(b"editable-docx-double")

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_selectable_text_output = Mock(
                side_effect=RuntimeError("可见文字 96/100 个非空字符")
            )

            app._run_batch(
                [source],
                converter.resolve_conversion_spec("pdf", "word"),
                None,
                converter.WORD_NATIVE,
                policy,
            )

            results, warnings = app._on_batch_complete.call_args.args
            self.assertEqual(results[0].status, "success")
            self.assertTrue(results[0].output.is_file())
            self.assertEqual(len(warnings), 1)
            self.assertIn("全文文字严格校验未完全通过", warnings[0])
            self.assertIn("96/100", warnings[0])
            app._on_batch_error.assert_not_called()

    def test_image_protected_policy_rejects_inconsistent_image_override(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "protected-text.pdf"
            source.write_bytes(b"text")
            policy = converter.PdfWordBatchPolicy(
                default_method="images",
                method_overrides={source: "images"},
                image_protected_paths={source},
            )
            self._bind_policy_snapshots(policy, source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app._create_converter = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )
            spec = converter.resolve_conversion_spec("pdf", "word")

            app._run_batch(
                [source],
                spec,
                None,
                "images",
                policy,
            )

            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertTrue(results[0].error)
            self.assertFalse((source.parent / "protected-text.docx").exists())
            app._create_converter.assert_not_called()
            app._on_batch_error.assert_not_called()

    def test_change_and_restore_during_preflight_still_converts_snapshot(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "change-and-restore.pdf"
            source.write_bytes(b"trusted")
            original_stat = source.stat()
            original_identity = converter._source_file_identity(source)

            app = self._app_for_pdf_fidelity_choice(converter.WORD_NATIVE)
            app.input_paths = [source]
            app._source_file_identity = converter._source_file_identity
            app._copy_verified_source_snapshot = (
                converter._copy_verified_source_snapshot
            )

            def analyze(snapshot, **_kwargs):
                snapshot_path = Path(snapshot)
                self.assertEqual(snapshot_path.read_bytes(), b"trusted")
                source.write_bytes(b"hostile")
                source.write_bytes(b"trusted")
                os.utime(
                    source,
                    ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns),
                )
                return self._fidelity_report(snapshot_path)

            with patch.object(
                converter,
                "analyze_pdf_fidelity_risk",
                side_effect=analyze,
            ):
                policy = app._choose_pdf_word_fidelity_mode(
                    converter.WORD_NATIVE
                )

            self.assertEqual(
                converter._source_file_identity(source),
                original_identity,
            )
            self.assertEqual(policy.source_snapshots[source].read_bytes(), b"trusted")

            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            converted_bytes = []

            def create(_spec, _method, allow_image_fallback=True):
                del allow_image_fallback

                def convert(snapshot, output, _progress):
                    converted_bytes.append(Path(snapshot).read_bytes())
                    Path(output).write_bytes(b"editable word")

                return convert

            app._create_converter = Mock(side_effect=create)
            spec = converter.resolve_conversion_spec("pdf", "word")
            app._run_batch(
                [source],
                spec,
                None,
                converter.WORD_NATIVE,
                policy,
            )

            self.assertEqual(converted_bytes, [b"trusted"])
            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(results[0].status, "success")
            self.assertTrue(results[0].output.is_file())
            app._on_batch_error.assert_not_called()

    def test_source_change_during_conversion_rejects_staged_output(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "changed-during-conversion.pdf"
            source.write_bytes(b"trusted")
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE
            )
            self._bind_policy_snapshots(policy, source)

            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, _method, allow_image_fallback=True):
                del allow_image_fallback

                def convert(snapshot, output, _progress):
                    self.assertEqual(Path(snapshot).read_bytes(), b"trusted")
                    Path(output).write_bytes(b"unverified word")
                    source.write_bytes(b"changed")

                return convert

            app._create_converter = Mock(side_effect=create)
            spec = converter.resolve_conversion_spec("pdf", "word")
            app._run_batch(
                [source],
                spec,
                None,
                converter.WORD_NATIVE,
                policy,
            )

            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertIn("转换期间发生变化", results[0].error)
            self.assertFalse(
                (source.parent / "changed-during-conversion.docx").exists()
            )
            app._on_batch_error.assert_not_called()

    def test_source_changed_after_preflight_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "changed-after-preflight.pdf"
            source.write_bytes(b"original")
            original_stat = source.stat()
            original_identity = converter._source_file_identity(source)
            policy = converter.PdfWordBatchPolicy(default_method="images")
            self._bind_policy_snapshots(policy, source)
            self.assertEqual(policy.source_identities[source], original_identity)
            snapshot_path = policy.source_snapshots[source]
            source.write_bytes(b"replaced")
            os.utime(
                source,
                ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns),
            )
            replacement_identity = converter._source_file_identity(source)
            self.assertEqual(original_identity[:4], replacement_identity[:4])
            self.assertNotEqual(original_identity[4], replacement_identity[4])
            self.assertEqual(snapshot_path.read_bytes(), b"original")

            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app._create_converter = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )
            spec = converter.resolve_conversion_spec("pdf", "word")

            app._run_batch([source], spec, None, "images", policy)

            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertTrue(results[0].error)
            self.assertFalse((source.parent / "changed-after-preflight.docx").exists())
            app._create_converter.assert_not_called()
            app._on_batch_error.assert_not_called()

    def test_pdf_to_word_without_preflight_policy_is_rejected(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._cancel_event = threading.Event()
        app._on_batch_error = Mock()
        app.root = SimpleNamespace(
            after=lambda _delay, callback, *args: callback(*args)
        )
        spec = converter.resolve_conversion_spec("pdf", "word")

        app._run_batch([], spec, None, converter.WORD_NATIVE, None)

        app._on_batch_error.assert_called_once()
        self.assertIn("缺少版式预检", app._on_batch_error.call_args.args[0])

    def test_editable_table_libreoffice_failure_never_offers_image_fallback(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = False
        app._ask_yes_no_from_worker = Mock(return_value=True)
        spec = converter.resolve_conversion_spec("pdf", "word")
        convert = app._create_converter(
            spec,
            "libreoffice",
            allow_image_fallback=False,
        )

        with patch.object(
            converter,
            "pdf_to_word_via_libreoffice",
            side_effect=RuntimeError("anonymous export failure"),
        ), patch.object(converter, "pdf_to_word_via_images") as image_convert:
            with self.assertRaisesRegex(RuntimeError, "未回退到.*图片"):
                convert("source.pdf", "output.docx", Mock())

        app._ask_yes_no_from_worker.assert_not_called()
        image_convert.assert_not_called()

    def test_editable_table_validation_rejects_any_missing_table_text(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=100,
            table_texts=("X" * 100,),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )
        incomplete_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=80,
            table_texts=("X" * 80,),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=incomplete_summary,
        ), self.assertRaisesRegex(RuntimeError, "80/100"):
            converter.ConverterApp._validate_editable_table_output(
                "missing-table-text.docx",
                report,
                Mock(),
            )
    def test_editable_table_validation_rejects_missing_text_outside_table(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((("A", "B"),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
            selectable_text="标题姓名AB",
            selectable_text_character_count=6,
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=1,
            cell_count=2,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((("A", "B"),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
            large_page_drawing_count=0,
            visible_text="标题AB",
            visible_text_character_count=4,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "全文可见文字 4/6"):
            converter.ConverterApp._validate_editable_table_output(
                "missing-outside-table-text.docx",
                report,
                Mock(),
            )

    def test_selectable_text_output_validation_is_exact(self):
        report = self._table_double(
            selectable_text="标题Alpha",
            selectable_text_character_count=7,
        )
        matching = self._table_double(
            visible_text="Alpha标题",
            visible_text_character_count=7,
        )
        progress = Mock()
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=matching,
        ):
            converter.ConverterApp._validate_selectable_text_output(
                "matching-text.docx",
                report,
                progress,
            )
        progress.assert_called_once()

        missing = self._table_double(
            visible_text="Alpha标",
            visible_text_character_count=6,
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=missing,
        ), self.assertRaisesRegex(RuntimeError, "6/7"):
            converter.ConverterApp._validate_selectable_text_output(
                "missing-text.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_requires_every_detected_table(self):
        report = self._table_double(
            table_count=2,
            table_text_character_count=100,
            table_texts=("A" * 50, "B" * 50),
            table_shapes=((2, 2), (2, 2)),
            table_cell_counts=(4, 4),
        )
        incomplete_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=100,
            table_texts=("A" * 50,),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=incomplete_summary,
        ), self.assertRaisesRegex(RuntimeError, "1/2"):
            converter.ConverterApp._validate_editable_table_output(
                "missing-table.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_compares_each_table_text_content(self):
        report = self._table_double(
            table_count=2,
            table_text_character_count=8,
            table_texts=("ABCD", "WXYZ"),
            table_shapes=((1, 2), (1, 2)),
            table_cell_counts=(2, 2),
        )
        unrelated_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=8,
            table_texts=("ABCD", "QQQQ"),
            table_shapes=((1, 2), (1, 2)),
            table_cell_counts=(2, 2),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=unrelated_summary,
        ), self.assertRaisesRegex(RuntimeError, "1/2"):
            converter.ConverterApp._validate_editable_table_output(
                "unrelated-table-text.docx",
                report,
                Mock(),
            )

        preserved_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=8,
            table_texts=("ABCD", "WXYZ"),
            table_shapes=((1, 2), (1, 2)),
            table_cell_counts=(2, 2),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=preserved_summary,
        ):
            converter.ConverterApp._validate_editable_table_output(
                "preserved-table-text.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_reversed_table_text(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=8,
            table_texts=("ABCDEFGH",),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
        )
        reversed_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=1,
            cell_count=1,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=8,
            table_texts=("HGFEDCBA",),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=reversed_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "reversed-table-text.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_collapsed_large_table(self):
        table_text = "SCORE" * 250
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((33, 26),),
            table_cell_counts=(515,),
        )
        collapsed_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=1,
            cell_count=1,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=collapsed_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "collapsed-table.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_gridspan_cell_collapse(self):
        table_text = "SCORE" * 250
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((33, 26),),
            table_cell_counts=(515,),
        )
        merged_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=33,
            cell_count=33,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((33, 26),),
            table_cell_counts=(33,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=merged_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "gridspan-collapse.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_any_grid_shape_reduction(self):
        table_text = "VALUE" * 250
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((33, 26),),
            table_cell_counts=(515,),
        )
        reduced_summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=27,
            cell_count=430,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((27, 21),),
            table_cell_counts=(430,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=reduced_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "reduced-grid-area.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_split_source_table(self):
        table_text = "".join(chr(0x4E00 + index) for index in range(100))
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((10, 5),),
            table_cell_counts=(50,),
        )
        split_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=10,
            cell_count=50,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text[:90], table_text[90:]),
            table_shapes=((9, 5), (1, 5)),
            table_cell_counts=(45, 5),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=split_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "split-table.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_multiple_small_fragments(self):
        table_text = "".join(chr(0x4E00 + index) for index in range(100))
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((10, 10),),
            table_cell_counts=(100,),
        )
        split_summary = self._table_double(
            has_editable_table=True,
            table_count=5,
            row_count=12,
            cell_count=100,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(
                table_text[:80],
                table_text[80:85],
                table_text[85:90],
                table_text[90:95],
                table_text[95:],
            ),
            table_shapes=((8, 10), (1, 5), (1, 5), (1, 5), (1, 5)),
            table_cell_counts=(80, 5, 5, 5, 5),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=split_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "fragmented-table.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_short_table_fragment(self):
        table_text = "ABCDEFGHIJ"
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((5, 2),),
            table_cell_counts=(10,),
        )
        split_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=5,
            cell_count=10,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text[:8], table_text[8:]),
            table_shapes=((4, 2), (1, 2)),
            table_cell_counts=(8, 2),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=split_summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "short-fragmented-table.docx",
                report,
                Mock(),
            )

    def test_table_text_normalization_preserves_meaningful_identity(self):
        for source, output in (("1 0", "10"), ("①", "1"), ("Ａ", "A")):
            with self.subTest(source=source, output=output):
                self.assertNotEqual(
                    converter._normalize_table_text(source),
                    converter._normalize_table_text(output),
                )
                self.assertFalse(
                    converter._table_cell_matrix_preserved(
                        ((source,),),
                        ((output,),),
                        (1, 1),
                        (1, 1),
                        source_full_width_span_rows=(False,),
                    )
                )

    def test_table_text_normalization_allows_only_layout_equivalence(self):
        self.assertEqual(
            converter._normalize_table_text("（6月）"),
            converter._normalize_table_text("(6月)"),
        )
        self.assertEqual(
            converter._normalize_table_text("中 文"),
            converter._normalize_table_text("中文"),
        )
        self.assertEqual(
            converter._normalize_table_text("A B"),
            "A B",
        )
    def test_structure_gate_rejects_fragment_before_main_table(self):
        preserved, matched, compared = converter._table_structures_preserved(
            ("ABCDEFGHIJ",),
            ((5, 2),),
            (10,),
            ("IJ", "ABCDEFGH"),
            ((1, 2), (4, 2)),
            (2, 8),
        )

        self.assertFalse(preserved)
        self.assertEqual((matched, compared), (0, 1))

    def test_structure_gate_aggregates_multiple_short_fragments(self):
        preserved, matched, compared = converter._table_structures_preserved(
            ("ABCDEFGHIJ",),
            ((5, 2),),
            (10,),
            ("ABCDEFGH", "I", "J"),
            ((4, 2), (1, 1), (1, 1)),
            (8, 1, 1),
        )

        self.assertFalse(preserved)
        self.assertEqual((matched, compared), (0, 1))

    def test_editable_table_validation_rejects_unrelated_layout_table(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=8,
            table_texts=("ABCDEFGH",),
            table_shapes=((4, 2),),
            table_cell_counts=(8,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=5,
            cell_count=10,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=16,
            table_texts=("ABCDEFGH", "WORKITEM"),
            table_shapes=((4, 2), (1, 2)),
            table_cell_counts=(8, 2),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "2/1"):
            converter.ConverterApp._validate_editable_table_output(
                "extra-layout-table.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_matches_structure_in_order(self):
        report = self._table_double(
            table_count=2,
            table_text_character_count=16,
            table_texts=("ABCDEFGH", "QRSTUVWX"),
            table_shapes=((10, 5), (4, 3)),
            table_cell_counts=(50, 12),
        )
        valid_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=14,
            cell_count=62,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=16,
            table_texts=("ABCDEFGH", "QRSTUVWX"),
            table_shapes=((10, 5), (4, 3)),
            table_cell_counts=(50, 12),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=valid_summary,
        ):
            converter.ConverterApp._validate_editable_table_output(
                "ordered-tables.docx",
                report,
                Mock(),
            )

        extra_summary = self._table_double(
            has_editable_table=True,
            table_count=3,
            row_count=15,
            cell_count=64,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=22,
            table_texts=("LAYOUT", "ABCDEFGH", "QRSTUVWX"),
            table_shapes=((1, 2), (10, 5), (4, 3)),
            table_cell_counts=(2, 50, 12),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=extra_summary,
        ), self.assertRaisesRegex(RuntimeError, "3/2"):
            converter.ConverterApp._validate_editable_table_output(
                "extra-tables.docx",
                report,
                Mock(),
            )

        reordered_summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=14,
            cell_count=62,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=16,
            table_texts=("QRSTUVWX", "ABCDEFGH"),
            table_shapes=((4, 3), (10, 5)),
            table_cell_counts=(12, 50),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=reordered_summary,
        ), self.assertRaisesRegex(RuntimeError, "1/2"):
            converter.ConverterApp._validate_editable_table_output(
                "reordered-tables.docx",
                report,
                Mock(),
            )
    def test_editable_table_validation_rejects_content_concentrated_in_top_left_cell(
        self,
    ):
        report = self._table_double(
            table_count=1,
            table_text_character_count=4,
            table_texts=("ABCD",),
            table_cell_matrices=((("A", "B"), ("C", "D")),),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=4,
            table_texts=("ABCD",),
            table_cell_matrices=((("ABCD", ""), ("", "")),),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "top-left.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_ten_by_ten_reduced_to_eight_rows(
        self,
    ):
        table_text = "VALUE" * 20
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((10, 10),),
            table_cell_counts=(100,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=8,
            cell_count=80,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_shapes=((8, 10),),
            table_cell_counts=(80,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "reduced-rows.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_swapped_cell_values(self):
        source_text = "NameScoreAlice95"
        output_text = "ScoreNameAlice95"
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(source_text),
            table_texts=(source_text,),
            table_cell_matrices=(
                (("Name", "Score"), ("Alice", "95")),
            ),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(output_text),
            table_texts=(output_text,),
            table_cell_matrices=(
                (("Score", "Name"), ("Alice", "95")),
            ),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "swapped.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_missing_checkmark(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=3,
            table_texts=("☑同意",),
            table_cell_matrices=((("☑", "同意"),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=1,
            cell_count=2,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=2,
            table_texts=("同意",),
            table_cell_matrices=((("", "同意"),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "2/3"):
            converter.ConverterApp._validate_editable_table_output(
                "missing-checkmark.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_merged_row_text_moved_within_row(
        self,
    ):
        table_text = "ABCDHeading"
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_cell_matrices=(
                (
                    ("A", "B", "C", "D"),
                    ("Heading", None, None, None),
                ),
            ),
            table_shapes=((2, 4),),
            table_full_width_span_rows=((False, True),),
            table_cell_counts=(5,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=8,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_cell_matrices=(
                (
                    ("A", "B", "C", "D"),
                    ("", "", "Heading", ""),
                ),
            ),
            table_shapes=((2, 4),),
            table_cell_counts=(8,),
        )
        progress = Mock()

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "merged-row.docx",
                report,
                progress,
            )

        progress.assert_not_called()
    def test_editable_table_validation_fails_closed_without_shapes(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=8,
            table_texts=("ABCDEFGH",),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=8,
            table_texts=("ABCDEFGH",),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaises(RuntimeError):
            converter.ConverterApp._validate_editable_table_output(
                "missing-shapes.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_fails_closed_on_incomplete_source_records(self):
        report = self._table_double(
            table_count=2,
            table_text_character_count=8,
            table_texts=("ABCD",),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
        )
        summary = self._table_double(
            has_editable_table=True,
            table_count=2,
            row_count=2,
            cell_count=2,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=8,
            table_texts=("ABCD", "WXYZ"),
            table_shapes=((1, 1), (1, 1)),
            table_cell_counts=(1, 1),
        )
        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaises(RuntimeError):
            converter.ConverterApp._validate_editable_table_output(
                "incomplete-source-records.docx",
                report,
                Mock(),
            )
    def test_editable_table_validation_rejects_locked_content_control(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=20,
        )
        locked_summary = self._table_double(
            has_editable_table=False,
            table_count=1,
            row_count=1,
            cell_count=1,
            document_protected=False,
            locked_content_control_table_count=1,
            table_text_character_count=20,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=locked_summary,
        ), self.assertRaisesRegex(RuntimeError, "内容控件锁定"):
            converter.ConverterApp._validate_editable_table_output(
                "locked-table.docx",
                report,
                Mock(),
            )

    def test_validation_failure_removes_output_and_batch_continues(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            table_source = folder / "editable-table.pdf"
            plain_source = folder / "plain.pdf"
            table_source.write_bytes(b"table")
            plain_source.write_bytes(b"plain")
            policy = converter.PdfWordBatchPolicy(
                default_method=converter.WORD_NATIVE,
                method_overrides={table_source: converter.WORD_NATIVE},
                editable_table_reports={
                    table_source: SimpleNamespace(
                        table_text_character_count=20
                    )
                },
            )
            self._bind_policy_snapshots(policy, table_source, plain_source)
            app = converter.ConverterApp.__new__(converter.ConverterApp)
            app._cancel_event = threading.Event()
            app._progress_update = Mock()
            app._status_update = Mock()
            app._on_batch_complete = Mock()
            app._on_batch_error = Mock()
            app.root = SimpleNamespace(
                after=lambda _delay, callback, *args: callback(*args)
            )

            def create(_spec, method, allow_image_fallback=True):
                del allow_image_fallback

                def convert(_source, output, _progress):
                    Path(output).write_bytes(method.encode("ascii"))

                return convert

            app._create_converter = Mock(side_effect=create)
            app._validate_editable_table_output = Mock(
                side_effect=RuntimeError("editable table validation failed")
            )
            spec = converter.resolve_conversion_spec("pdf", "word")
            app._run_batch(
                [table_source, plain_source],
                spec,
                None,
                converter.WORD_NATIVE,
                policy,
            )

            results = app._on_batch_complete.call_args.args[0]
            self.assertEqual(
                [result.status for result in results],
                ["failed", "success"],
            )
            self.assertFalse((folder / "editable-table.docx").exists())
            self.assertTrue((folder / "plain.docx").is_file())
            app._on_batch_error.assert_not_called()
    def test_confirmed_blank_table_never_uses_manual_image_mode(self):
        source = Path("anonymous-blank-table.pdf")
        app = self._app_for_pdf_fidelity_choice("images")
        app.input_paths = [source]
        app._method_availability.side_effect = (
            lambda method: method in {"images", converter.WORD_NATIVE}
        )
        app._ask_ok_cancel_from_worker.return_value = True
        report = self._fidelity_report(
            source,
            is_complex=True,
            editable_table_candidate=True,
            table_count=1,
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            table_texts=("",),
            table_cell_matrices=((('', ''), ('', '')),),
            table_full_width_span_rows=((False, False),),
            table_layout_suspected=True,
            reasons=("检测到表格结构",),
        )

        with patch.object(
            converter,
            "analyze_pdf_fidelity_risk",
            return_value=report,
        ):
            policy = app._choose_pdf_word_fidelity_mode("images")

        self.assertEqual(policy.method_overrides[source], converter.WORD_NATIVE)
        self.assertEqual(policy.editable_table_reports[source], report)
        self.assertEqual(policy.image_protected_paths, {source})
        app._ask_ok_cancel_from_worker.assert_called_once()

    def test_blank_source_table_accepts_matching_empty_word_table_structure(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=0,
            table_texts=("",),
            table_cell_matrices=((('', ''), ('', '')),),
            table_shapes=((2, 2),),
            table_full_width_span_rows=((False, False),),
            table_cell_counts=(4,),
        )
        summary = self._table_double(
            has_editable_table=False,
            has_editable_table_structure=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=0,
            table_texts=("",),
            table_cell_matrices=((('', ''), ('', '')),),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            large_page_drawing_count=0,
        )
        progress = Mock()

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ):
            converter.ConverterApp._validate_editable_table_output(
                "blank-table.docx",
                report,
                progress,
            )

        progress.assert_called_once()

    def test_vertical_merge_continuation_cannot_move_text_to_another_column(self):
        table_text = "GroupHeaderValue"
        report = self._table_double(
            table_count=1,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_cell_matrices=(
                (("Group", "Header"), (None, "Value")),
            ),
            table_shapes=((2, 2),),
            table_full_width_span_rows=((False, False),),
            table_cell_counts=(3,),
            table_cell_spans=(
                (
                    (0, 0, 2, 1),
                    (0, 1, 1, 1),
                    (1, 1, 1, 1),
                ),
            ),
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=2,
            cell_count=4,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=len(table_text),
            table_texts=(table_text,),
            table_cell_matrices=(
                (("Group", "Header"), ("Value", "")),
            ),
            table_shapes=((2, 2),),
            table_cell_counts=(4,),
            large_page_drawing_count=0,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "0/1"):
            converter.ConverterApp._validate_editable_table_output(
                "wrong-column.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_accepts_complete_border_evidence(self):
        border_edges = (
            (
                (True, True, False, True),
                (True, False, True, True),
            ),
        )
        report = self._table_double(
            table_count=1,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((('A', 'B'),),),
            table_shapes=((1, 2),),
            table_full_width_span_rows=((False,),),
            table_cell_counts=(2,),
            table_cell_border_edges=border_edges,
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=1,
            cell_count=2,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((('A', 'B'),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
            table_cell_border_edges=border_edges,
            large_page_drawing_count=0,
        )
        progress = Mock()

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ):
            converter.ConverterApp._validate_editable_table_output(
                "matching-borders.docx",
                report,
                progress,
            )

        progress.assert_called_once()

    def test_editable_table_validation_rejects_missing_cell_border(self):
        source_edges = (((True, True, True, True),),)
        output_edges = (((False, True, True, True),),)
        report = self._table_double(
            table_count=1,
            table_text_character_count=1,
            table_texts=("A",),
            table_cell_matrices=((('A',),),),
            table_shapes=((1, 1),),
            table_full_width_span_rows=((False,),),
            table_cell_counts=(1,),
            table_cell_border_edges=source_edges,
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=1,
            cell_count=1,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=1,
            table_texts=("A",),
            table_cell_matrices=((('A',),),),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
            table_cell_border_edges=output_edges,
            large_page_drawing_count=0,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "单元格边框与源 PDF 不一致"):
            converter.ConverterApp._validate_editable_table_output(
                "missing-border.docx",
                report,
                Mock(),
            )

    def test_editable_table_validation_rejects_extra_cell_border(self):
        source_edges = (((False, True, True, True),),)
        output_edges = (((True, True, True, True),),)
        report = self._table_double(
            table_count=1,
            table_text_character_count=1,
            table_texts=("A",),
            table_cell_matrices=((('A',),),),
            table_shapes=((1, 1),),
            table_full_width_span_rows=((False,),),
            table_cell_counts=(1,),
            table_cell_border_edges=source_edges,
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=1,
            cell_count=1,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=1,
            table_texts=("A",),
            table_cell_matrices=((('A',),),),
            table_shapes=((1, 1),),
            table_cell_counts=(1,),
            table_cell_border_edges=output_edges,
            large_page_drawing_count=0,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "单元格边框与源 PDF 不一致"):
            converter.ConverterApp._validate_editable_table_output(
                "extra-border.docx",
                report,
                Mock(),
            )

    def test_large_page_image_rejects_otherwise_matching_editable_table(self):
        report = self._table_double(
            table_count=1,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((('A', 'B'),),),
            table_shapes=((1, 2),),
            table_full_width_span_rows=((False,),),
            table_cell_counts=(2,),
        )
        summary = self._table_double(
            has_editable_table=True,
            has_editable_table_structure=True,
            table_count=1,
            row_count=1,
            cell_count=2,
            document_protected=False,
            locked_content_control_table_count=0,
            table_text_character_count=2,
            table_texts=("AB",),
            table_cell_matrices=((('A', 'B'),),),
            table_shapes=((1, 2),),
            table_cell_counts=(2,),
            large_page_drawing_count=1,
        )

        with patch.object(
            converter,
            "inspect_editable_docx_tables",
            return_value=summary,
        ), self.assertRaisesRegex(RuntimeError, "大面积页面图片"):
            converter.ConverterApp._validate_editable_table_output(
                "photo-backed-table.docx",
                report,
                Mock(),
            )


if __name__ == "__main__":
    unittest.main()
