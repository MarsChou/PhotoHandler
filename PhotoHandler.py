# coding=utf-8
import os
import exifread
import datetime,time
import pytz
import shutil
from win32com.propsys import propsys, pscon

def TimeStampToTime(timestamp):
    timeStruct = time.localtime(timestamp)
    return time.strftime('%Y%m%d_%H%M%S',timeStruct)

def getExif(full_file_name):
    fd = open(full_file_name, 'rb')
    tags = exifread.process_file(fd)
    fd.close()
    return tags

img_folder = "E:\\微云照片备份\\"
new_folder = "E:\\WeiYunBackup\\"
if __name__ == '__main__':
    g = os.walk(img_folder)
    for path, d, file_list in g:
        for file_name in file_list:
            full_file_name = os.path.join(path, file_name)
            new_full_file_name = full_file_name
            file_suffix = os.path.splitext(file_name)[-1]
            new_name = ""
            if file_suffix.lower() in ['.mp4', '.mov', '.avi', ".jpg", ".png"]:
                if file_suffix.lower() in ['.mp4', '.mov', '.avi']:
                        # 如果是视频，使用媒体创建日期
                        properties = propsys.SHGetPropertyStoreFromParsingName(full_file_name)
                        dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
                        properties = None # release file handle
                        if dt:
                            if not isinstance(dt, datetime.datetime):
                                # In Python 2, PyWin32 returns a custom time type instead of
                                # using a datetime subclass. It has a Format method for strftime
                                # style formatting, but let's just convert it to datetime:
                                dt = datetime.datetime.fromtimestamp(int(dt))
                                dt = dt.replace(tzinfo=pytz.timezone('UTC'))
                            dt_shanghai = dt.astimezone(pytz.timezone('Asia/Shanghai'))
                            new_name = dt_shanghai.strftime('%Y%m%d_%H%M%S')
                else:
                    tags = getExif(full_file_name)
                    FIELD = 'EXIF DateTimeOriginal'
                    if FIELD in tags:
                        # 使用拍摄日期
                        new_name = str(tags[FIELD]).replace(':', '').replace(' ', '_')
                    else:    
                        # 使用文件修改时间
                        new_name = TimeStampToTime(os.stat(full_file_name).st_mtime)
            if not new_name:
                print("no new_name")
                print(full_file_name)
                dest_folder = new_folder + "Others"
                if not os.path.exists(dest_folder):
                    os.mkdir(dest_folder)
                new_full_file_name =  os.path.join(dest_folder, file_name)
            else:
                dest_folder = new_folder + new_name.split("_")[0][:6]
                if not os.path.exists(dest_folder):
                    os.mkdir(dest_folder)
                new_full_file_name = os.path.join(dest_folder, new_name + file_suffix)
            
            if not os.path.exists(new_full_file_name):
                shutil.copyfile(full_file_name, new_full_file_name)