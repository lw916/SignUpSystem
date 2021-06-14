import os

# 获取当前地址：
BasePath = os.getcwd()

# 网页地址
Web_Url = ""

# 当前签到状态：
SignupStatus = 0
# 输出文件名：
ExcelFileName = "SignUp.xlxs"
# 上传文件地址(py根目录下upload文件夹）
UploadDir = "./upload/"
# 保存文件地址(py根目录下result文件夹）
SaveDir = "./result/"

# 读取excel文件临时数据
FileTemp = []
# 写入excel临时数据
WriteTemp = []

# 后台服务器信息：
Port = 8000
Host = "0.0.0.0"

# 二维码地址
Url = "http://192.168.1.75:8000/home#/sign"
