import subprocess
import unittest
from types import SimpleNamespace

from engine_models import ConversionBackend, EngineState
from macos_office import (
    AUTOMATION_SETTINGS,
    OSASCRIPT_PATH,
    SOURCE_ALREADY_OPEN_DETAIL,
    SOURCE_ALREADY_OPEN_MARKER,
    MacOSOfficeError,
    MacOSPowerPointBackend,
    MacOSWordBackend,
)


def completed(returncode=0, stdout="", stderr=""):
    return SimpleNamespace(returncode=returncode, stdout=stdout, stderr=stderr)


class RecordingRunner:
    def __init__(self, result=None, error=None):
        self.result = result or completed()
        self.error = error
        self.calls = []

    def __call__(self, command, **kwargs):
        self.calls.append((command, kwargs))
        if self.error is not None:
            raise self.error
        return self.result


class MacOSOfficeProbeTests(unittest.TestCase):
    def test_shallow_probe_reports_missing_without_starting_osascript(self):
        runner = RecordingRunner()
        backend = MacOSWordBackend(installed_checker=lambda: False, runner=runner)

        status = backend.probe()

        self.assertIs(status.state, EngineState.MISSING)
        self.assertFalse(status.installed)
        self.assertEqual(runner.calls, [])

    def test_shallow_probe_reports_installed_as_unverified(self):
        runner = RecordingRunner()
        backend = MacOSPowerPointBackend(
            installed_checker=lambda: True,
            runner=runner,
        )

        status = backend.probe(deep=False)

        self.assertIs(status.state, EngineState.UNVERIFIED)
        self.assertTrue(status.installed)
        self.assertFalse(status.usable)
        self.assertEqual(runner.calls, [])

    def test_deep_probe_uses_argument_list_and_shell_false(self):
        runner = RecordingRunner(completed(stdout="16.89\n"))
        backend = MacOSWordBackend(installed_checker=lambda: True, runner=runner)

        status = backend.probe(deep=True)

        self.assertIs(status.state, EngineState.AVAILABLE)
        self.assertEqual(status.detail, "16.89")
        command, kwargs = runner.calls[0]
        self.assertIsInstance(command, list)
        self.assertEqual(command[:2], [OSASCRIPT_PATH, "-e"])
        self.assertEqual(command[3:], [])
        self.assertFalse(kwargs["shell"])

    def test_permission_denial_and_timeout_are_structured_probe_results(self):
        permission_runner = RecordingRunner(
            completed(1, stderr="Not authorized to send Apple events. (-1743)")
        )
        permission_backend = MacOSWordBackend(
            installed_checker=lambda: True,
            runner=permission_runner,
        )
        timeout_backend = MacOSPowerPointBackend(
            installed_checker=lambda: True,
            runner=RecordingRunner(
                error=subprocess.TimeoutExpired([OSASCRIPT_PATH], timeout=20)
            ),
        )

        self.assertIs(
            permission_backend.probe(deep=True).state,
            EngineState.PERMISSION_DENIED,
        )
        self.assertIs(timeout_backend.probe(deep=True).state, EngineState.TIMEOUT)

    def test_other_nonzero_probe_result_is_launch_failed(self):
        backend = MacOSWordBackend(
            installed_checker=lambda: True,
            runner=RecordingRunner(completed(1, stderr="Application is not running")),
        )

        status = backend.probe(deep=True)

        self.assertIs(status.state, EngineState.LAUNCH_FAILED)
        self.assertIn("Application is not running", status.detail)



class MacOSOfficeConversionTests(unittest.TestCase):
    def test_backends_implement_the_shared_contract_and_directions(self):
        word = MacOSWordBackend(installed_checker=lambda: True)
        powerpoint = MacOSPowerPointBackend(installed_checker=lambda: True)

        self.assertIsInstance(word, ConversionBackend)
        self.assertIsInstance(powerpoint, ConversionBackend)
        self.assertEqual(word.directions, frozenset({"word_to_pdf"}))
        self.assertEqual(
            powerpoint.directions,
            frozenset({"powerpoint_to_pdf"}),
        )

    def test_word_conversion_passes_unicode_paths_only_through_argv(self):
        runner = RecordingRunner()
        backend = MacOSWordBackend(installed_checker=lambda: True, runner=runner)
        source = "/Users/测试 文件/报告.docx"
        output = "/Users/测试 文件/报告 输出.pdf"
        progress = []

        backend.convert(source, output, lambda message, pct: progress.append((message, pct)))

        command, kwargs = runner.calls[0]
        script = command[2]
        self.assertEqual(command[3:], [source, output])
        self.assertIn("on run argv", script)
        self.assertNotIn(source, script)
        self.assertNotIn(output, script)
        self.assertIn("set sourceAlias to sourceFile as alias", script)
        self.assertIn("repeat with existingDocument in documents", script)
        self.assertIn("(full name of existingDocument) as text", script)
        self.assertIn("if existingFullName is sourcePath then", script)
        self.assertIn("set existingAlias to existingFullName as alias", script)
        self.assertIn("if existingAlias is sourceAlias then", script)
        self.assertIn(SOURCE_ALREADY_OPEN_MARKER, script)
        guard_index = script.index(SOURCE_ALREADY_OPEN_MARKER)
        self.assertLess(guard_index, script.index("set openedDocument to open sourceFile"))
        self.assertLess(guard_index, script.index("save as openedDocument"))
        self.assertIn("close openedDocument", script)
        self.assertLess(guard_index, script.index("close openedDocument"))
        self.assertNotIn("quit", script.casefold())
        self.assertFalse(kwargs["shell"])
        self.assertEqual(progress[-1], ("完成", 100))

    def test_powerpoint_conversion_is_safe_and_successful(self):
        runner = RecordingRunner()
        backend = MacOSPowerPointBackend(
            installed_checker=lambda: True,
            runner=runner,
        )
        source = "/Users/example/中文 幻灯片.pptx"
        output = "/Users/example/中文 幻灯片.pdf"

        backend.convert(source, output, lambda *_args: None)

        command, kwargs = runner.calls[0]
        script = command[2]
        self.assertEqual(command[0], OSASCRIPT_PATH)
        self.assertEqual(command[3:], [source, output])
        self.assertIn("on run argv", script)
        self.assertNotIn(source, script)
        self.assertIn("set sourceAlias to sourceFile as alias", script)
        self.assertIn("repeat with existingPresentation in presentations", script)
        self.assertIn("(full name of existingPresentation) as text", script)
        self.assertIn("if existingFullName is sourcePath then", script)
        self.assertIn("set existingAlias to existingFullName as alias", script)
        self.assertIn("if existingAlias is sourceAlias then", script)
        self.assertIn(SOURCE_ALREADY_OPEN_MARKER, script)
        guard_index = script.index(SOURCE_ALREADY_OPEN_MARKER)
        self.assertIn("set openedPresentation to open sourceFile", script)
        self.assertLess(guard_index, script.index("set openedPresentation to open sourceFile"))
        self.assertNotIn("active presentation", script.casefold())
        self.assertLess(guard_index, script.index("save openedPresentation"))
        self.assertIn("close openedPresentation", script)
        self.assertLess(guard_index, script.index("close openedPresentation"))
        close_lines = [
            line.strip() for line in script.splitlines() if line.strip().startswith("close ")
        ]
        self.assertTrue(all(line == "close openedPresentation" for line in close_lines))
        self.assertNotIn("quit", script.casefold())
        self.assertFalse(kwargs["shell"])

    def test_permission_failure_raises_readable_error_with_settings_path(self):
        runner = RecordingRunner(
            completed(1, stderr="automation denied: AppleEvent error -1743")
        )
        backend = MacOSWordBackend(installed_checker=lambda: True, runner=runner)

        with self.assertRaises(MacOSOfficeError) as context:
            backend.convert("/tmp/source.docx", "/tmp/output.pdf", lambda *_args: None)

        self.assertIs(context.exception.status.state, EngineState.PERMISSION_DENIED)
        self.assertIn(AUTOMATION_SETTINGS, str(context.exception))

    def test_timeout_error_suggests_retry_without_automation_settings(self):
        timeout = subprocess.TimeoutExpired([OSASCRIPT_PATH], timeout=600)
        backend = MacOSPowerPointBackend(
            installed_checker=lambda: True,
            runner=RecordingRunner(error=timeout),
        )

        with self.assertRaises(MacOSOfficeError) as context:
            backend.convert("/tmp/source.pptx", "/tmp/output.pdf", lambda *_args: None)

        self.assertIs(context.exception.status.state, EngineState.TIMEOUT)
        self.assertNotIn(AUTOMATION_SETTINGS, str(context.exception))
        self.assertIn("关闭", str(context.exception))

    def test_missing_error_suggests_installation_without_automation_settings(self):
        backend = MacOSWordBackend(installed_checker=lambda: False)

        with self.assertRaises(MacOSOfficeError) as context:
            backend.convert("/tmp/source.docx", "/tmp/output.pdf", lambda *_args: None)

        message = str(context.exception)
        self.assertIs(context.exception.status.state, EngineState.MISSING)
        self.assertIn("官方下载页", message)
        self.assertNotIn(AUTOMATION_SETTINGS, message)

    def test_launch_failed_error_suggests_application_and_document_checks(self):
        backend = MacOSWordBackend(
            installed_checker=lambda: True,
            runner=RecordingRunner(completed(1, stderr="Application is not running")),
        )

        with self.assertRaises(MacOSOfficeError) as context:
            backend.convert("/tmp/source.docx", "/tmp/output.pdf", lambda *_args: None)

        message = str(context.exception)
        self.assertIs(context.exception.status.state, EngineState.LAUNCH_FAILED)
        self.assertIn("正常启动", message)
        self.assertIn("文档可正常打开", message)
        self.assertNotIn(AUTOMATION_SETTINGS, message)

    def test_already_open_source_is_launch_failed_with_protective_message(self):
        runner = RecordingRunner(
            completed(
                1,
                stderr=f"execution error: {SOURCE_ALREADY_OPEN_MARKER} (-17001)",
            )
        )
        backend = MacOSPowerPointBackend(
            installed_checker=lambda: True,
            runner=runner,
        )

        with self.assertRaises(MacOSOfficeError) as context:
            backend.convert("/tmp/source.pptx", "/tmp/output.pdf", lambda *_args: None)

        self.assertIs(context.exception.status.state, EngineState.LAUNCH_FAILED)
        self.assertEqual(context.exception.status.detail, SOURCE_ALREADY_OPEN_DETAIL)
        message = str(context.exception)
        self.assertIn("未开始转换", message)
        self.assertIn("不会保存或关闭它", message)
        self.assertIn("请先关闭该源文件后重试", message)
        self.assertNotIn(AUTOMATION_SETTINGS, message)


if __name__ == "__main__":
    unittest.main()
