class EtreeMixin(object):

    IGNORED_FIELDS=[
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR"
        ,"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr"
        ]

    def filter_item(self, item):
        return item[0] not in self.IGNORED_FIELDS

    def assert_equal_tree(self, lhs, rhs):
        """
        Compares two instances of ElementTree are equivalent.
        """
        self.assertEqual(lhs.tag, rhs.tag)
        self.assertEqual(len(lhs), len(rhs))
        self.assertEqual(lhs.text or '', rhs.text or '')
        self.assertEqual(sorted(filter(self.filter_item, lhs.items())), sorted(filter(self.filter_item, rhs.items())))
        for lhs_child, rhs_child in zip(lhs, rhs):
            self.assert_equal_tree(lhs_child, rhs_child)


def get_document_body_part(document):
    for part in document.parts.values():
        if part.getroot().tag.endswith('}document'):
            return part

    raise AssertionError("main document body not found in document.parts")
