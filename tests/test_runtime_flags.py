import os
import unittest
from unittest.mock import patch

from app.services import runtime_flags


class RuntimeFlagsTests(unittest.TestCase):
    def test_is_desktop_api_sync_enabled_respects_env_override(self):
        with patch.dict(os.environ, {"ROTA_DESKTOP_SYNC_API": "1"}, clear=False):
            self.assertTrue(runtime_flags.is_desktop_api_sync_enabled())
        with patch.dict(os.environ, {"ROTA_DESKTOP_SYNC_API": "0"}, clear=False):
            self.assertFalse(runtime_flags.is_desktop_api_sync_enabled())

    def test_can_read_from_api_requires_sync_enabled(self):
        with patch.dict(
            os.environ,
            {
                "ROTA_DESKTOP_SYNC_API": "0",
                "ROTA_ALLOW_REMOTE_READ": "1",
                "ROTA_SOURCE_OF_TRUTH": "api-central",
            },
            clear=False,
        ):
            self.assertFalse(runtime_flags.can_read_from_api())

    def test_can_read_from_api_true_when_all_flags_allow(self):
        with patch.dict(
            os.environ,
            {
                "ROTA_DESKTOP_SYNC_API": "1",
                "ROTA_ALLOW_REMOTE_READ": "1",
                "ROTA_SOURCE_OF_TRUTH": "api-central",
            },
            clear=False,
        ):
            self.assertTrue(runtime_flags.can_read_from_api())

    def test_can_read_from_api_false_for_local_source_of_truth(self):
        with patch.dict(
            os.environ,
            {
                "ROTA_DESKTOP_SYNC_API": "1",
                "ROTA_ALLOW_REMOTE_READ": "1",
                "ROTA_SOURCE_OF_TRUTH": "sqlite-local",
            },
            clear=False,
        ):
            self.assertFalse(runtime_flags.can_read_from_api())


if __name__ == "__main__":
    unittest.main()
