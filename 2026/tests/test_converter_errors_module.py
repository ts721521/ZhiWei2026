import unittest


class ConverterErrorsModuleSplitTests(unittest.TestCase):
    def test_converter_errors_module_exists_and_classifies_permission(self):
        from converter.errors import ConversionErrorType, classify_conversion_error

        result = classify_conversion_error(PermissionError("access denied"))
        self.assertEqual(ConversionErrorType.PERMISSION_DENIED, result.get("error_type"))
        self.assertFalse(result.get("is_retryable"))

    def test_office_converter_re_exports_error_classifier(self):
        from converter.errors import classify_conversion_error as split_classifier
        from office_converter import classify_conversion_error as office_classifier

        self.assertIs(split_classifier, office_classifier)


if __name__ == "__main__":
    unittest.main()
