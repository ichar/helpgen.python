# -*- coding: cp1251 -*-

import os
import sys
import datetime
import shutil

version_info = { 
    'version' : '2.03 beta', 
    'date'    : '2014-09-30', 
    'now'     : datetime.datetime.now().strftime('%Y-%m-%d'), 
    'time'    : datetime.datetime.now().strftime('%H:%M'), 'author' : 'ichar' 
}
version = 'version %(version)s, %(date)s, %(author)s' % version_info
short_version = 'version %(version)s � DoorHan' % version_info
date_version = '%(now)s %(time)s' % version_info

DATE_STAMP = '%Y%m%d'
DATETIME_STAMP = '%Y-%m-%d %H:%M:%S'

DS = "\\"
default_path = 'configurators'
default_images = ':\images'
service_files_list = ( \
    'buttons.txt', 'errors.xml', 'messages.xml', 
    'functions.txt', 'functions.js', 'functions.min.js',
    'configurators_comm.xml', 'configurators_local.xml', 'valid-construct.xml', 'local-valid-construct.xml',
    'constants.xml', 'styles_1.xml', 
    'calculationHelper.swf',
    'm2.cfg',
)
special_file_list = ( \
    'local-valid-construct.xml',
)

# Журнал изменений:
# -----------------
# 2.03: папка общих изображений: <root>\images
# 2.02: контроль комплектации изделия
# 2.01: контроль комплектации
# 2.00: html5 setup
# 1.12: параметр <modified> сделан необязательным
# 1.11: source final slash control
# 1.10: версия, date
# 1.09: кодировка файла coding: cp1251

def chdir(s):
    drive, path = os.path.splitdrive(s)
    p = ''
    for x in path.split(DS):
        if not x:
            continue
        p = os.path.join(p, x)
        v = os.path.join(drive, DS, p)
        if not os.path.exists(v):
            os.mkdir(v)
            #print '--> %s created' % v

def check_default(p, item):
    #print '--> %s:%s' % (p, item)
    return os.path.isfile(p) and (default_images in p or item in service_files_list) and 1 or 0 #default_path in item or 

def check_path(p, m, model=None, force=None):
    subdirs = p.split(DS)
    item = subdirs[-1]
    if model:
        if force:
            if model in subdirs or (check_default(p, item) and not '\\dus' in p):
                return 1
            else:
                return 0
        else:
            if default_path in subdirs and not (model in subdirs or check_default(p, item)):
                return 0
    modified = os.path.getmtime(p)
    D = datetime.datetime.fromtimestamp(modified)
    IsMatched = 0
    if D > m or (model in subdirs and item in special_file_list):
        #print '--> %s %s' % (D.strftime(DATETIME_STAMP), p)
        IsMatched = 1
    return IsMatched

def walk(source, destination, modified, model=None, force=None):
    n = 0
    for name in sorted(os.listdir(source)):
        p = os.path.join(source, name)
        IsMatched = check_path(p, modified, model=model, force=force)
        if os.path.isdir(p):
            n += walk(p, destination, modified, model, force)
        elif os.path.isfile(p):
            if IsMatched:
                drive, path = os.path.splitdrive(p)
                d = os.path.dirname('%s\\%s' % (destination, path))
                #print d, drive, destination, path
                chdir(d)
                try:
                    shutil.copy2(p, d)
                    print '--> %s copied into %s' % (p, d)
                except:
                    print '--> %s error copying into %s' % (p, d)
                n += 1
    return n


if __name__ == "__main__":
    argv = sys.argv

    if len(argv) < 4 or argv[1].lower() in ('/h', '/help', '-h', 'help', '--help'):
        print '--> DoorHan Inc.'
        print '--> *Web-Helper* configurator\'s setup utility.'
        print '--> '
        print '--> Format: setup.py <source> <destination> {<modified>|*} [<DUS>] [<force>]'
        print '--> '
        print '--> Where:'
        print '--> '
        print '-->   <source> - path to the source directory (server)'
        print '-->   <destination> - path to the destination directory (distibutive)'
        print '-->   <modified> - modification datetime as YYYYMMDD'
        print '-->   <DUS> - model name (xxx - for core changes only)'
        print '-->   <force> - Y/N (all files)'
        print '--> '
        print '--> %s[Python2]' % version

    else:
        source = argv[1]
        if source[-1] not in (DS, "/"):
            source += DS
        print '--> Source: %s' % source

        subdir = argv[3] in '*xX' and datetime.datetime.now().strftime(DATE_STAMP) or argv[3]

        chdir(os.path.join(argv[2], subdir))
        destination = os.path.join(argv[2], subdir, 'helper')
        chdir(destination)
        print '--> Destination: %s' % destination

        modified = datetime.datetime.strptime(subdir, DATE_STAMP)
        print '--> Modified: %s' % modified

        model = len(argv) > 4 and argv[4].lower() or None
        if model:
            print '--> Model: %s' % model

        force = len(argv) > 5 and argv[5].lower() == 'y' and 'Y'
        if force:
            print '--> Force: %s' % force

        n = walk(source, destination, modified, model=model, force=force)

        print '--> Total: %d' % n
