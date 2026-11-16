import unittest
from unittest.mock import patch, MagicMock
from app.version import parse_version, is_newer_version, APP_VERSION
from app.core.updater import AppUpdater

class TestUpdater(unittest.TestCase):
    def test_version_parsing(self):
        self.assertEqual(parse_version("1.0.0"), (1, 0, 0))
        self.assertEqual(parse_version("v1.2.3"), (1, 2, 3))
        self.assertEqual(parse_version("2.1"), (2, 1, 0))
        self.assertEqual(parse_version("v3"), (3, 0, 0))

    def test_is_newer_version(self):
        self.assertTrue(is_newer_version("v1.1.0", "1.0.0"))
        self.assertTrue(is_newer_version("2.0.0", "1.9.9"))
        self.assertTrue(is_newer_version("v1.0.1", "1.0.0"))
        self.assertFalse(is_newer_version("1.0.0", "1.0.0"))
        self.assertFalse(is_newer_version("v0.9.9", "1.0.0"))
        self.assertFalse(is_newer_version("1.0.0", "1.1.0"))

    @patch("requests.get")
    def test_check_for_update_has_newer(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.json.return_value = {
            "tag_name": "v1.1.0",
            "name": "AssignmentFixer v1.1.0",
            "body": "- Added auto-updater\n- Improved UI",
            "assets": [
                {
                    "name": "AssignmentFixer-v1.1.0-windows.zip",
                    "browser_download_url": "https://github.com/test/download.zip",
                    "size": 52428800
                }
            ]
        }
        mock_get.return_value = mock_resp

        updater = AppUpdater(current_version="1.0.0")
        res = updater.check_for_update()

        self.assertIsNotNone(res)
        self.assertEqual(res["tag_name"], "v1.1.0")
        self.assertEqual(res["download_url"], "https://github.com/test/download.zip")
        self.assertEqual(res["asset_name"], "AssignmentFixer-v1.1.0-windows.zip")
        self.assertEqual(res["asset_size"], 52428800)

    @patch("requests.get")
    def test_check_for_update_same_version(self, mock_get):
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.json.return_value = {
            "tag_name": "v1.0.0",
            "assets": []
        }
        mock_get.return_value = mock_resp

        updater = AppUpdater(current_version="1.0.0")
        res = updater.check_for_update()
        self.assertIsNone(res)

if __name__ == "__main__":
    unittest.main()
