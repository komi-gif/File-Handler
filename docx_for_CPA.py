# -*-coding: utf-8 -*-
from win32com.client import Dispatch
import os

word_path = 'F:\\Study\\1 Exams\\1 CPA\\2 讲义\\2022年讲义\\财务成本管理\\1 基础知识\\基础精讲班-闫华红(笔记版word)\\'
word = Dispatch('Word.Application')
word.Visible = 0
file_name = os.listdir(word_path)
level1 = ["第一节","第二节","第三节","第四节","第五节","第六节","第七节","第八节","第九节",]
# level1 = ["第一部分","第二部分","第三部分","第四部分","第五部分","第六部分","第七部分","第八部分","第九部分","第十部分",]
# level2 = ["考点1","考点2","考点3","考点4","考点5","考点6","考点7","考点8","考点9","考点10","考点11",]
# level2 = ["【知识点1】","【知识点2】","【知识点3】","【知识点4】","【知识点5】","【知识点6】","【知识点7】","【知识点8】",]
level2 = ["一、","二、","三、","四、","五、","六、","七、","八、","九、","十、","十一、","十二、","十三、","十四、",]
level3 = ["（一）","（二）","（三）","（四）","（五）","（六）","（七）","（八）","（九）","（十）",]
# level4 = ["1. ","2. ","3. ","4. ","5. ","6. ","7. ","8. ","9. ","10. ","11. ",]
for name in file_name:

    word.Visible = 0
    doc = word.Documents.Open(word_path+name)

    for i in level1:
        search_range = doc.Range()
        for m in range(5):
            search_range.Find.Execute(FindText=i)
            search_range.Paragraphs(1).Style = doc.Styles(-2)
    for j in level2:
        search_range = doc.Range()
        for i in range(10):
            search_range.Find.Execute(FindText=j)
            search_range.Paragraphs(1).Style = doc.Styles(-3)
    for k in level3:
        search_range = doc.Range()
        for i in range(15):
            search_range.Find.Execute(FindText=k)
            search_range.Paragraphs(1).Style = doc.Styles(-4)
    # for l in level4:
    #     search_range = doc.Range()
    #     for i in range(10):
    #         search_range.Find.Execute(FindText=l)
    #         search_range.Paragraphs(1).Style = doc.Styles(-5)
    doc.Save()
    doc.Close()
