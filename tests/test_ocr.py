from unittest import TestCase

from ocr import parse_json_response


class OCRTests(TestCase):
    def test_json_in_codeblock(self):
        result = parse_json_response('''```json
{"supplier":"Тест", "invoice_number":"5", "invoice_date":"19.11.2026", "items": [
{"name":"Товар", "unit":"шт", "quantity":2, "unit_price":10.5, "amount":21}], "warnings":[]}
```''')
        self.assertEqual(len(result["items"]), 1)
        self.assertEqual(result["items"][0]["amount"], 21)

    def test_invalid_response(self):
        with self.assertRaises(ValueError):
            parse_json_response("не json")
