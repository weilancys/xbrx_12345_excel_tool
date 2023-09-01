import unittest
import datetime
import os
from xbrx_12345_excel_tool import excel


class TestZhengfuBiao(unittest.TestCase):
    def setUp(self):
        source_file = os.path.join(os.path.dirname(__file__), "templates", "zhengfubiao_example.xls")
        self.zhengfubiao = excel.ZhengFuBiao(source_file)


    def test_all_rows_loaded(self):
        self.assertEqual(len(self.zhengfubiao.rows), 333) # the example file has 332 valid data rows


    def test_get_rows_by_time(self):
        start_time = datetime.datetime(2020, 11, 27, 16, 6, 3) # 2020/11/27 16:06:03
        end_time = datetime.datetime(2020, 11, 27, 16, 11, 38) # 2020/11/27 16:11:38
        rows = self.zhengfubiao.get_rows_by_time(start_time, end_time)

        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0][0], "201127151804745594")

    
    def test_get_earliest_row(self):
        earliest_row = self.zhengfubiao.get_earliest_row()
        self.assertEqual(earliest_row[0], "201119161421360124")


    def test_get_latest_row(self):
        latest_row = self.zhengfubiao.get_latest_row()
        self.assertEqual(latest_row[0], "201127145829183624")

    
    def test_get_all_ids(self):
        ids = self.zhengfubiao.get_all_ids()
        self.assertEqual(len(ids), 333)
        self.assertEqual(ids[0], "201127145829183624")
        self.assertEqual(ids[-1], "201119161421360124")

    
    def test_include_duplicates_works(self):
        ids_with_duplicates = self.zhengfubiao.get_all_ids(include_duplicates=True)
        ids_without_duplicates = self.zhengfubiao.get_all_ids(include_duplicates=False)
        self.assertNotEqual(len(ids_with_duplicates), len(ids_without_duplicates))


class TestSangaoTemplate(unittest.TestCase):
    def setUp(self):
        zhengfubiao_file = os.path.join(os.path.dirname(__file__), "templates", "zhengfubiao_example.xls")
        self.zhengfubiao = excel.ZhengFuBiao(zhengfubiao_file)


    def test_save(self):
        rows = self.zhengfubiao.rows[:10]
        sangao_template = excel.SangaoTemplate(rows)
        save_path =  "test_output.xlsx"
        sangao_template.save(save_path)
        self.assertTrue(os.path.exists(save_path))


    def test_generate_template(self):
        start_time = datetime.datetime(2020, 11, 27, 16, 6, 3) # 2020/11/27 16:06:03
        end_time = datetime.datetime(2020, 11, 27, 16, 11, 38) # 2020/11/27 16:11:38
        rows = self.zhengfubiao.get_rows_by_time(start_time, end_time)
        sangao_template = excel.SangaoTemplate(rows)
        save_path = "generated_output.xlsx"
        sangao_template.save(save_path)
        self.assertTrue(os.path.exists(save_path))

    
    def tearDown(self):
        if os.path.exists("test_output.xlsx"):
            os.unlink("test_output.xlsx")

        if os.path.exists("generated_output.xlsx"):
            os.unlink("generated_output.xlsx")



class TestSangaoBiao(unittest.TestCase):
    def setUp(self):
        sangaobiao_file = os.path.join(os.path.dirname(__file__), "templates", "sangaobiao_example.xlsx")
        self.sangaobiao = excel.SangaoBiao(sangaobiao_file)
    

    def test_all_data_rows_loaded(self):
        self.assertEqual(self.sangaobiao.row_count, 4174)

    
    def test_get_all_12345_ids(self):
        ids = self.sangaobiao.get_all_12345_ids()
        self.assertEqual(ids[0], "201127105122067200")
        self.assertEqual(ids[-1], "201124104927405348")

    
    def test_include_duplicates_works(self):
        ids_with_duplicates = self.sangaobiao.get_all_12345_ids()
        ids_without_duplicates = self.sangaobiao.get_all_12345_ids(include_duplicates=False)
        self.assertNotEqual(len(ids_with_duplicates), len(ids_without_duplicates))

    
    def test_get_12345_ids_histogram(self):
        histogram = self.sangaobiao.get_12345_ids_histogram()
        self.assertIsInstance(histogram, dict)
        self.assertEqual(len(histogram), len(self.sangaobiao.get_all_12345_ids(include_duplicates=False)))
        self.assertNotIn(0, histogram.values())
        self.assertIn(1, histogram.values())

    
    def test_get_recurrent_rows(self):
        recurrent_rows = self.sangaobiao.get_recurrent_rows()
        self.assertNotEqual(recurrent_rows, [])

    
    def test_get_12345_ids_histogram(self):
        histogram = self.sangaobiao.get_12345_ids_histogram(recurrent_id_only=True)
        self.assertIsInstance(histogram, dict)
        self.assertEqual(len(histogram), 527)


class TestValidationReport(unittest.TestCase):
    def setUp(self):
        sangaobiao_file = os.path.join(os.path.dirname(__file__), "templates", "sangaobiao_example.xlsx")
        self.sangaobiao = excel.SangaoBiao(sangaobiao_file)
        zhengfubiao_file = os.path.join(os.path.dirname(__file__), "templates", "zhengfubiao_example.xls")
        self.zhengfubiao = excel.ZhengFuBiao(zhengfubiao_file)

    
    def test_generate_report_text(self):
        validation_report = excel.ValidationReport(self.zhengfubiao, self.sangaobiao)
        report_text = validation_report.generate_report_text()
        self.assertIn("!DOCTYPE html", report_text)
    

    def test_missing_ids(self):
        validation_report = excel.ValidationReport(self.zhengfubiao, self.sangaobiao)
        self.assertEqual(validation_report.has_missing_ids, True)
        self.assertNotEqual(validation_report.missing_ids, [])

