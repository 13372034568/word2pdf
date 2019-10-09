import os
import time

from win32com.client import Dispatch


# word.application
# kwps.application
def g_pdf(names_specials=[], max_cnt=-1, app_name='kwps.application', dir_name_docx=None, dir_name_pdf=None,
          overwrite=False):
    index = 0
    if not os.path.exists(dir_name_pdf):
        os.makedirs(dir_name_pdf)
    for i in os.listdir(dir_name_docx):
        if not i.endswith('.doc') and not i.endswith('.docx'):
            continue
        fn = os.path.splitext(i)[0]
        fp_docx = os.path.join(os.path.abspath(dir_name_docx), i)
        fp_pdf = os.path.join(os.path.abspath(dir_name_pdf), fn + '.pdf')
        if (names_specials and fn not in names_specials) \
                or (0 < max_cnt <= index):
            continue
        if not overwrite and os.path.exists(fp_pdf):
            index += 1
            print(index)
            print('%s exists, skip...' % fp_pdf)
            continue
        try:
            index += 1
            print(index)
            print('process %s' % fp_docx)
            w = Dispatch(app_name)
            doc = w.Documents.Open(fp_docx)
            doc.ExportAsFixedFormat(fp_pdf, 17)
            w.Quit()
            time.sleep(1)
        except:
            print('process %s failed' % (fn))


if __name__ == '__main__':
    dir_name_docx = os.path.abspath('word')
    dir_name_pdf = os.path.abspath('pdf')
    g_pdf(app_name='kwps.application',
          dir_name_docx=dir_name_docx, dir_name_pdf=dir_name_pdf, overwrite=True)
    # g_pdf(app_name='word.application',
    #       dir_name_docx=dir_name_docx, dir_name_pdf=dir_name_pdf, overwrite=True)
