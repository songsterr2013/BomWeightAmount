import traceback
import util
from BomWeightAmount import BomWeightAmount

logger = util.get_logger(__file__)
s, config = util.load_config()
if __name__ == '__main__':
    s, config = util.load_config()

    if s:

        try:
            bom_path = config["bom_path"]
            bom_list_file = config["bom_list_file"]

            # Instantiate
            bwa = BomWeightAmount(bom_path, bom_list_file)

            # 重整格式
            bwa.reformatting()

            # 從EXCEL中取出每一個BOM的(第一個字, BOM檔名)
            # 取得BOM裡面的內容
            # 寫到原本的EXCEL BOM LIST中
            for prefix, file_name, weight_loc, area_loc in bwa.parse_bom():
                print(prefix, file_name, weight_loc, area_loc)
                bwa.get_bom_content(prefix, file_name)
                bwa.write_data(weight_loc, area_loc)

            bwa.save()

        except SystemError as e:
            logger.error("錯誤:{}".format(traceback.format_exc()))

    else:
        logger.error("沒有設定檔或設定檔異常：{}".format(traceback.format_exc()))