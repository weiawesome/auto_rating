from utils import *

######################################################################
# 建立 reference_files 資料夾 放入 index.html 與 name_list.xlsx
# 建立 result 資料夾 以儲存結果
######################################################################

num = 5
penalty = 5
score = [30, 50, 70, 80, 90, 100]

idScoreDf = get_id_score_pd()
nameListDf = get_name_id_df()

resultDf = merge_result(nameListDf, idScoreDf)

save_to_excel(resultDf)
