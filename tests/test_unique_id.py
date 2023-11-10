import unittest

from mailmerge import UniqueIdsManager

class UniqueIdsManagerTest(unittest.TestCase):
    """
    Testing UniqueIdsManager class
    """

    def test_unique_id_manager_register_id(self):
        """
        Tests if the next record field works
        """
        tests = [
            ("id", 2, None),
            ("id", 2, 3),
            ("id", None, 4),
            ("footer", 1, None),
            ("footer", None, 2)
        ]
        id_man = UniqueIdsManager()

        for type_id, obj_id, new_id in tests:
            self.assertEqual(id_man.register_id(type_id, obj_id=obj_id), new_id)

        self.assertEqual(id_man.register_id_str("footer2"), "footer3")
