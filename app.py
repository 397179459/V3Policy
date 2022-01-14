import copy
import logging
from flask import Flask, render_template, request, send_from_directory
import os
import datetime

logger = logging.getLogger()
# logger.setLevel(logging.INFO)
logger.setLevel(logging.DEBUG)

ABS_PATH = os.path.abspath('.')
TARGET_ZIP = ABS_PATH + r'\PolicyFile\ZipSuccessPolicy'

ALLOW_EXTENSIONS = set(['xlsx', 'xls', 'csv'])

app = Flask(__name__)


# app.config['UPLOAD_DIR'] = UPLOAD_DIR
# app.config['SERVER_DIR'] = SERVER_DIR
# app.config['TARGET_DIR'] = TARGET_DIR
# app.config['TARGET_ZIP'] = TARGET_ZIP
# app.config['ALLOW_EXTENSIONS'] = ALLOW_EXTENSIONS


def allowed_file(filename):
    if '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOW_EXTENSIONS:
        return True
    else:
        return False


@app.route('/')
def index():
    # logging.debug("this is debug")
    # logging.info("this is info")
    # logging.warning("this is warning")
    # logging.error("this is error")
    # logging.critical("this is critical")
    return render_template("upload_file.html")


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        global now
        global FIRST_DIR, SERVER_DIR, TARGET_DIR, TARGET_ZIP

        now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')

        FIRST_DIR = ABS_PATH + r'\PolicyFile\{0}'.format(now)
        # UPLOAD_DIR = FIRST_DIR + '\SourcePolicy'
        SERVER_DIR = FIRST_DIR + '\ServerPolicy'
        TARGET_DIR = FIRST_DIR + '\TargetPolicy'
        # TARGET_ZIP = FIRST_DIR + '\ZipPolicy'

        os.mkdir(FIRST_DIR)
        # os.mkdir(UPLOAD_DIR)
        os.mkdir(SERVER_DIR)
        os.mkdir(TARGET_DIR)
        # os.mkdir(TARGET_ZIP)

        logging.debug('文件上传成功')

        global FACTORY
        FACTORY = request.form['RadioFactory']
        # logging.debug(FACTORY)

        global PRODUCTSPEC
        PRODUCTSPEC = request.form['ProductSpec']

        file = request.files['file']
        if file and allowed_file(file.filename):
            # sourceFileNamePre = file.filename.split('.')[0]
            serverFileNamePre = 'PolicyUpload'
            filenamesuf = file.filename.split('.')[1]

            # sourceFileName = sourceFileNamePre + now + '.' + filenamesuf

            global tfilename
            tfilename = serverFileNamePre + now + '.' + filenamesuf

            # saveSourceFileName = os.path.join(app.config['UPLOAD_DIR'], sourceFileName)
            saveFileName = os.path.join(SERVER_DIR, tfilename)

            # f1 = copy.copy(file)
            # f2 = file
            file.save(saveFileName)
            # f1.save(saveSourceFileName)

            try:
                import TPFOParse

                zipFile = TPFOParse.main()

            except:
                return render_template('fileError.html')

            return render_template('fileDownload.html', tfileName=zipFile)

        else:
            return render_template('fileError.html')


@app.route('/download/<ttfilename>')
def download(ttfilename):
    dir = os.path.join(TARGET_ZIP)
    logging.debug(dir)
    logging.debug(ttfilename)
    return send_from_directory(directory=dir, path=ttfilename, as_attachment=True)


if __name__ == '__main__':
    app.run()
