import unittest
from pathlib import Path

from engine_models import ConversionBackend, EngineState, EngineStatus


class EngineStatusTests(unittest.TestCase):
    def test_all_required_states_are_defined(self):
        self.assertEqual(
            {state.value for state in EngineState},
            {
                "checking",
                "available",
                "missing",
                "unverified",
                "permission_denied",
                "launch_failed",
                "timeout",
                "unsupported",
            },
        )

    def test_status_properties_distinguish_presence_and_usability(self):
        expected = {
            EngineState.CHECKING: (False, False, False),
            EngineState.AVAILABLE: (True, True, True),
            EngineState.MISSING: (False, False, True),
            EngineState.UNVERIFIED: (True, False, True),
            EngineState.PERMISSION_DENIED: (True, False, True),
            EngineState.LAUNCH_FAILED: (True, False, True),
            EngineState.TIMEOUT: (True, False, True),
            EngineState.UNSUPPORTED: (False, False, True),
        }
        for state, properties in expected.items():
            with self.subTest(state=state):
                status = EngineStatus(state, "detail")
                self.assertEqual(
                    (status.installed, status.usable, status.complete),
                    properties,
                )

    def test_string_state_is_normalized(self):
        status = EngineStatus("available", None)
        self.assertIs(status.state, EngineState.AVAILABLE)
        self.assertEqual(status.detail, "")

    def test_invalid_state_is_rejected(self):
        with self.assertRaises(ValueError):
            EngineStatus("broken")


class FakeBackend:
    id = "fake"
    display_name = "Fake backend"
    directions = frozenset({"pdf_to_word"})

    def probe(self, deep=False):
        return EngineStatus(EngineState.AVAILABLE if deep else EngineState.UNVERIFIED)

    def convert(self, source, output, progress):
        progress("done", 100)
        Path(output).write_bytes(Path(source).read_bytes())


class ConversionBackendTests(unittest.TestCase):
    def test_runtime_protocol_accepts_structural_backend(self):
        backend = FakeBackend()
        self.assertIsInstance(backend, ConversionBackend)
        self.assertEqual(backend.probe().state, EngineState.UNVERIFIED)
        self.assertEqual(backend.probe(deep=True).state, EngineState.AVAILABLE)


if __name__ == "__main__":
    unittest.main()
