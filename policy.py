import logging
from flask import Flask, render_template, request, send_from_directory
import os
import datetime

global logLevel
logLevel = logging.DEBUG

# log，测试和生产log打印配置不同
logger = logging.getLogger()
# logger.setLevel(logging.INFO)
logger.setLevel(logLevel)

ABS_PATH = os.path.abspath('.')
# 输出的zip文件夹，有一个show file界面需要展示这个目录
TARGET_ZIP = ABS_PATH + r'/PolicyFile/ZipSuccessPolicy'

# 上传选择文件，限制的文件后缀
ALLOW_EXTENSIONS = set(['xlsx', 'xls', 'csv'])

app = Flask(__name__)


# app.config['UPLOAD_DIR'] = UPLOAD_DIR
# app.config['SERVER_DIR'] = SERVER_DIR
# app.config['TARGET_DIR'] = TARGET_DIR
# app.config['TARGET_ZIP'] = TARGET_ZIP
# app.config['ALLOW_EXTENSIONS'] = ALLOW_EXTENSIONS

# check 上传的文件是否是在要求的文件类型中
def allowed_file(filename):
    if '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOW_EXTENSIONS:
        return True
    else:
        return False


# 主页面
@app.route('/')
def index():
    return render_template("upload_file.html")


# form 提交之后处理，以及重定向
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        global now
        global FIRST_DIR, SERVER_DIR, TARGET_DIR, TARGET_ZIP

        # 当前时间戳，文件名需要用
        now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')

        # 先创建时间戳文件夹，，再依次创建policy原档，和转换之后存的位置，后期删除和归档方便
        FIRST_DIR = ABS_PATH + r'/PolicyFile/{0}'.format(now)
        # UPLOAD_DIR = FIRST_DIR + '\SourcePolicy'
        SERVER_DIR = FIRST_DIR + '/ServerPolicy'
        TARGET_DIR = FIRST_DIR + '/TargetPolicy'
        # TARGET_ZIP = FIRST_DIR + '\ZipPolicy'

        os.mkdir(FIRST_DIR)
        # os.mkdir(UPLOAD_DIR)
        os.mkdir(SERVER_DIR)
        os.mkdir(TARGET_DIR)
        # os.mkdir(TARGET_ZIP)

        logging.debug('文件上传成功')

        # Factory和 productSpec 都是前端传的
        global FACTORY
        FACTORY = request.form['RadioFactory']

        global PRODUCTSPEC
        PRODUCTSPEC = request.form['ProductSpec']

        file = request.files['file']
        if file and allowed_file(file.filename):
            # sourceFileNamePre = file.filename.split('.')[0]
            # 主要害怕文件名乱码，所以是重命名存在服务器
            serverFileNamePre = 'PolicyUpload'
            filenamesuf = file.filename.split('.')[1]

            # sourceFileName = sourceFileNamePre + now + '.' + filenamesuf

            # 服务器上完整的文件名
            global tfilename
            tfilename = serverFileNamePre + now + '.' + filenamesuf

            # saveSourceFileName = os.path.join(app.config['UPLOAD_DIR'], sourceFileName)
            saveFileName = os.path.join(SERVER_DIR, tfilename)

            # 保存文件
            file.save(saveFileName)

            try:
                # 只能在这里import，如果在开头import，会导致循环引用报错
                import TPFOParse

                oldPath = os.path.join(SERVER_DIR, tfilename)

                # 开始处理原始文件
                zipFile = TPFOParse.main(oldPath)

            except Exception as e:
                msg = repr(e)
                logging.debug(msg)
                # 处理失败
                return render_template('fileError.html', errorMsg=msg)

            # 处理成功，返回到成功可下载页面
            return render_template('fileDownload.html', zipfileName=zipFile)

        else:
            return render_template('fileError.html')


# 下载地址
@app.route('/download/<ttfilename>')
def download(ttfilename):
    dir = os.path.join(TARGET_ZIP)
    logging.debug(dir)
    logging.debug(ttfilename)
    return send_from_directory(directory=dir, path=ttfilename, as_attachment=True)


# 展示zip dir中的所有文件
@app.route('/files')
def files():
    files = []
    for filename in os.listdir(TARGET_ZIP):
        path = os.path.join(TARGET_ZIP, filename)
        if os.path.isfile(path):
            files.append(filename)

    return render_template('showfiles.html', files=files)


# download all files function
@app.route("/files/<path:path>")
def get_file(path):
    """Download a file."""
    return send_from_directory(TARGET_ZIP, path, as_attachment=True)


if __name__ == '__main__':
    app.run()
