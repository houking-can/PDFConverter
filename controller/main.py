import os
from subprocess import Popen, PIPE, STDOUT
import time
import argparse
from os.path import exists, join
import threading
import shutil


def iter_files(path):
    """Walk through all files located under a root path."""
    if os.path.isfile(path):
        yield path
    elif os.path.isdir(path):
        for dir_path, _, file_names in os.walk(path):
            for f in file_names:
                yield os.path.join(dir_path, f)
    else:
        raise RuntimeError('Path %s is invalid' % path)



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='bat run of the pdf2html')

    parser.add_argument('--input', '-i', required=True, action='store',
                        help='Path of files you want to convert')
    parser.add_argument('--output', '-o', default=join(os.getcwd(), 'output'), action='store',
                        help='Path to store converted files')
    parser.add_argument('--oldpath', required=True ,help='Path to store old pdfs')
    parser.add_argument('--exe', '-e', required=True, action='store',
                        help='Path of pdf2html.exe')
    parser.add_argument('--replace', '-r', default=False, action='store',
                        help='Whether replace exist files')
    parser.add_argument('--timeout', '-t', type=float, default=20, action='store',
                        help='Timeout for a single convert')
    parser.add_argument('--format', '-f', default='html', action='store',
                        help='Format you want to convert: '
                             'xml, txt, doc, docx, '
                             'ps, jpeg, jpe, jpg, jpf, '
                             'jpx, j2k, j2c, jpc, rtf, '
                             'accesstext, tif, tiff')
    args = parser.parse_args()
    main(args)
