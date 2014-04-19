class EtreeMixin(object):
    def assert_equal_tree(self, lhs, rhs):
        """
        Compares two instances of ElementTree are equivalent.
        """
        self.assertEqual(lhs.tag, rhs.tag)
        self.assertEqual(len(lhs), len(rhs))
        self.assertEqual(lhs.text, rhs.text)
        self.assertEqual(sorted(lhs.items()), sorted(rhs.items()))
        for lhs_child, rhs_child in zip(lhs, rhs):
            self.assert_equal_tree(lhs_child, rhs_child)
