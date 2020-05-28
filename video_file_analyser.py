import os
import pwd
import grp
from pathlib import Path
from openpyxl import Workbook
from pymediainfo import MediaInfo
import stat
import subprocess
import xml.etree.ElementTree as ET

curr_dir = os.path.dirname(os.path.realpath(__file__))
print(f'curr_dir: {curr_dir}')

wb = Workbook()
ws = wb.active
ws.title = 'videos'
video_file_extensions = ['.mkv', '.mp4', '.m4p', '.m4v', '.avi', '.webm', '.flv', '.vob',
                         '.ogv', '.ogg', '.gif', '.mov', '.wmv', '.mpg', '.mpeg', '.3gp', ]
ws.append(['type', 'dir', 'video file', 'user', 'group', 'size', 'mode', 'width', 'height', 'frame rate', 'codec', 'encoded lib name'])
found_video_files = {}
# for dirpath, dirnames, filenames in os.walk(Path(curr_dir, 'movies')):
for dirpath, dirnames, filenames in os.walk(curr_dir):
    video_files = filenames  # [f for f in filenames if f[f.rfind('.'):].lower() in video_file_extensions]
    vd = Path(dirpath)
    st = vd.stat()
    permission = st.st_mode
    vd_mode = permission & 0o777
    dir_type = permission & 0o777000
    vd_uid = vd.owner()  # pwd.getpwuid(st.st_uid)
    vd_gid = vd.group()  # grp.getgrgid(st.st_gid)
    ws.append(['d', str(Path(dirpath).relative_to(curr_dir)), str(Path(dirpath).relative_to(curr_dir)), vd_uid, vd_gid, '', f'{vd_mode:o}'])
    for video_file in video_files:
        # tmp = []
        # vf = os.path.join(dirpath, video_file)
        vf = Path(dirpath, video_file)
        file_size = os.path.getsize(vf)
        st = vf.stat()
        permission = st.st_mode
        mode = permission & 0o777
        file_type = permission & 0o777000
        uid = vf.owner()  # pwd.getpwuid(st.st_uid)
        gid = vf.group()  # grp.getgrgid(st.st_gid)
        # print(f'owner: {uid} - group: {gid} - file_type: {file_type:o} - mode: {mode:o} - video_file: {vf}')
        tmp = ['f', str(Path(dirpath).relative_to(curr_dir)), video_file, uid, gid, file_size, f'{mode:o}']

        # video_file_info = {'dirpath': dirpath, 'file': video_file, 'file_type': file_type, 'file_mode': mode, }
        media_info = MediaInfo.parse(vf)
        for track in media_info.tracks:
            # print(track.track_type)
            if track.track_type == 'Video':
                tmp += [track.width, track.height, track.frame_rate, track.codec, track.encoded_library_name]
        ws.append(tmp)

        # o = subprocess.check_output(['mediainfo', '--Output=XML', vf])
        # # tree = ET.parse('country_data.xml')
        # # root = tree.getroot()
        # root = ET.fromstring(o.encode('utf-8'))
        # for child in root:
        #     print(child.tag, child.attrib)
        # for neighbor in root.iter():
        #     print(neighbor.attrib)
        # # media = root.find('media')
        # # print(f'media: {media}')
        # # for c in media.findall('track'):
        # #     print('found')
        # #     print(c.get('type'))

wb.save(Path(curr_dir, 'test.xlsx'))
