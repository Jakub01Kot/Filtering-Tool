import unittest
from openpyxl import Workbook
from io import BytesIO
from main import extract_info_from_excel  # Adjust the import as per your script's name

class TestLinkedInExtractor(unittest.TestCase):

    def setUp(self):
        # Creating a sample Excel workbook in memory
        self.wb = Workbook()
        ws = self.wb.active
        data = [
            ["Alice Adams"],
            ["1st"],
            ["Software Engineer"],
            ["California, USA"],
            ["Experience: 5 years"],
            ["Bob Brown"],
            ["2nd"],
            ["Data Scientist"],
            ["New York, USA"],
            ["Experience: 5 years"]

        ]
        for row in data:
            ws.append(row)

        # Saving workbook to a BytesIO object to simulate file read
        self.virtual_file = BytesIO()
        self.wb.save(self.virtual_file)
        self.virtual_file.seek(0)

    def test_extractor(self):
        df = extract_info_from_excel(self.virtual_file)  # Adjust the function to return df for this test

        # Check the names
        print(df)  # Print the first few rows of the DataFrame
        self.assertEqual(df["Name"].iloc[0], "Alice Adams")
        self.assertEqual(df.iloc[1]["Name"], "Bob Brown")

        # Check the professions
        self.assertEqual(df.iloc[0]["Profession"], "Software Engineer")
        self.assertEqual(df.iloc[1]["Profession"], "Data Scientist")


if __name__ == '__main__':
    unittest.main()
