# excel文档操作导入包
import xlrd
import xlwt
# 获取文件名包
from werkzeug.utils import secure_filename
# 常量
import config
# 生成二维码包
from MyQR import myqr
# 后台系统包
from flask import Flask, request, send_from_directory, render_template, redirect

'''excel文件操作部分'''


# 生成excel文件
def write_excel(data):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Sheet')
    for i in range(0, len(data)):
        worksheet.write(i, 0, label=data[i])
    workbook.save(config.SaveDir + "out.xls")


# 读取excel文档
def load_excel(path):
    # 打开excel文件
    workbook = xlrd.open_workbook(path)
    # 打开excel文件的sheet1
    content = workbook.sheet_by_name("Sheet")
    # 创建人员数组
    persons = []
    # 遍历excel文件
    for i in range(0, content.nrows):
        dict_temp = {"name": content.row_values(i)[0], "status": 0}
        persons.append(dict_temp)
    return persons


app = Flask(__name__, template_folder='dict', static_folder='dict/static')


# 防止跨域，请求头增加允许跨域
@app.after_request
def cors(environ):
    environ.headers['Access-Control-Allow-Origin'] = '*'
    environ.headers['Access-Control-Allow-Method'] = '*'
    environ.headers['Access-Control-Allow-Headers'] = 'x-requested-with,content-type'
    return environ


# 定位主页为登陆页
@app.route('/', methods=['GET'])
def hello():
    return redirect("/home")


# 渲染页面
@app.route('/home', methods=['GET'])
def template():
    return render_template('index.html')


# 签到系统登陆
@app.route('/login', methods=['GET'])
def login():
    #  获取get参数
    args = request.args
    if args is None:
        return {"status": 0, "msg": "登陆账户名或密码未填写"}
    # 获取账户名
    username = args.get("username")
    # 获取账户密码
    password = args.get("password")
    # 判断账户密码是否正确
    if username == "admin" and password == "admin":
        return {"status": 1, "msg": "欢迎使用签到系统"}
    else:
        return {"status": 0, "msg": "账户或密码错误"}


# 上传excel文件并获取签到名单及二维码图片
@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        try:
            # 获取上传文件的文件名
            filename = secure_filename(file.filename)
            # 保存文件
            file.save(config.UploadDir + filename)
            try:
                # 读取excel文件
                temp = load_excel(config.UploadDir + filename)
                # 读取数据存入临时文件
                config.FileTemp = temp
                # 获取需要签到人数
                SignNum = len(temp)
                # 生成二维码图片
                myqr.run(words=config.Url, save_dir="dict/static")
                return {"status": 1, "msg": "上传成功,数据已添加", "SignNum": SignNum, "UnSign": 0}
            except:
                return {"status": 0, "msg": "上传文件失败，可能是由于文件格式错误"}
        except:
            return {"status": 0, "msg": "未上传任何文件"}


# 停止签到：
@app.route('/stop', methods=['GET'])
def stop_sign():
    if len(config.FileTemp) == 0:
        return {"status": 0, "msg": "没有导入签到数据"}
    else:
        try:
            # 设置签到状态未不可签到
            config.SignupStatus = 0
            UnSign = []
            for i in range(0, len(config.FileTemp)):
                if config.FileTemp[i]['status'] == 0:
                    UnSign.append(config.FileTemp[i]['name'])
            config.WriteTemp = UnSign
            write_excel(config.WriteTemp)
            return {"status": 1, "msg": "停止签到成功", "url": "https://img01.yzcdn.cn/vant/cat.jpeg"}
        except:
            return {"status": 0, "msg": "停止签到失败"}


# 开始签到：
@app.route('/start', methods=['GET'])
def start_sign():
    if len(config.FileTemp) == 0:
        return {"status": 0, "msg": "没有导入签到数据"}
    else:
        try:
            # 设置签到状态未可签到
            config.SignupStatus = 1
            return {"status": 1, "msg": "开始签到成功", "url": "static/qrcode.png"}
        except:
            return {"status": 0, "msg": "开始签到失败"}


# 获取当前签到状态：
@app.route('/status', methods=['GET'])
def query_status():
    data = config.FileTemp
    if len(data) == 0:
        return {"status": 0, "msg": "没有上传数据"}
    else:
        # 未签到名单
        UnSign = []
        for i in range(0, len(data)):
            if data[i]['status'] == 0:
                # 获取未签到人名
                UnSign.append(data[i]['name'])
        return {"status": 1, "msg": "获取成功", "SignNum": len(data), "UnSign": len(UnSign), "list": UnSign}


# 签到
@app.route('/sign', methods=['GET'])
def sign():
    # 当签到状态为不可签到时
    if config.SignupStatus == 0:
        return {"status": 0, "msg": "超过签到时间或签到停止或未开始"}
    # 当签到状态为可签到时
    else:
        #  获取get参数
        args = request.args
        # 获取签到名
        name = args.get("name")
        # 获取账户密码
        status = args.get("status")
        SignStatus = 0
        for i in range(0, len(config.FileTemp)):
            if config.FileTemp[i]['name'] == name:
                if int(config.FileTemp[i]['status']) > 0:
                    return {"status": 0, "msg": "已签到"}
                else:
                    config.FileTemp[i]['status'] = status
                    SignStatus = 1
        if SignStatus == 0:
            return {"status": 0, "msg": "未找到人名记录"}
        else:
            return {"status": 1, "msg": "签到成功"}


# 签到excel文件下载
@app.route('/download')
def download():
    if len(config.WriteTemp) == 0:
        return {"status": 0, "msg": "未生成名单"}
    else:
        return send_from_directory(directory=config.SaveDir, path="out.xls", as_attachment=True)


# 后台启动函数
if __name__ == '__main__':
    app.config["SignupStatus"] = 0
    app.config['JSON_AS_ASCII'] = False
    app.run(port=config.Port, host=config.Host)
