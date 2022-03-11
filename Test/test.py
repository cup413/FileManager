import re


reg = '轮探1.*?录井总结报告'

s = 'D:\李晨星文件夹\项目文件\塔里木程小桂\data\完井报告\轮探1\轮探1地质录井总结报告1.docx'


a= re.search(reg, s)
print(a)
print(a.group())
print(re.search(reg, s))