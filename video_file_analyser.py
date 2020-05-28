#! /usr/bin/env python3

import os
import sys
import argparse
from pathlib import Path
from openpyxl import Workbook
from pymediainfo import MediaInfo


def validate_dir(text):
    path = Path(text).expanduser()
    if not path.exists():
        raise argparse.ArgumentTypeError(f"The path \"{path}\" does not exist!")
    return path


def parse_cli_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("path", help="Directory of video files.", type=validate_dir)
    parser.add_argument("--name", default="videos", help="name of output excel file.", type=str)
    # parser.add_argument("-v", "--verbose", help="verbose output.", action="store_true")

    return parser.parse_args()


def is_web(name):
    name = name.lower()
    return 'web-dl' in name or 'webrip' in name or '.web.' in name


def main():
    args = parse_cli_args()

    path = Path(args.path).absolute()
    print(f"Video Directory: '{path}'")

    wb = Workbook()
    ws = wb.active
    ws.title = 'videos'
    video_file_extensions = ['.mkv', '.mp4', '.m4p', '.m4v', '.avi', '.webm', '.flv', '.vob',
                             '.ogv', '.ogg', '.gif', '.mov', '.wmv', '.mpg', '.mpeg', '.3gp', ]

    ws.append(['type', 'dir', 'video file', 'ext', 'is web-dl', 'user', 'group', 'size', 'mode', 'width', 'height', 'frame rate', 'codec', 'encoded lib name', 'audio format', 'audio channels', 'audio bit rate', 'audio sampling rate', 'audio language'])
    for dirpath, dirnames, filenames in os.walk(path):
        dirpath = Path(dirpath)
        video_files = filenames  # [f for f in filenames if f[f.rfind('.'):].lower() in video_file_extensions]
        vid_dir = Path(dirpath)
        vid_dir_st = vid_dir.stat()
        permission = vid_dir_st.st_mode
        vd_mode = permission & 0o777
        dir_type = permission & 0o777000
        vid_dir_uid = vid_dir.owner()
        vid_dir_gid = vid_dir.group()
        ws.append(['d', str(dirpath.relative_to(path)), str(dirpath.relative_to(path)), '', '', vid_dir_uid, vid_dir_gid, '', f'{vd_mode:o}'])
        for video_file in video_files:
            vf = dirpath / video_file
            st = vf.stat()
            file_size = st.st_size
            permission = st.st_mode
            mode = permission & 0o777
            file_type = permission & 0o777000
            uid = vf.owner()
            gid = vf.group()
            tmp = ['f', str(dirpath.relative_to(path)), video_file, vf.suffix, 'X' if is_web(video_file) else '', uid, gid, file_size, f'{mode:o}']

            media_info = MediaInfo.parse(vf)
            for track in media_info.tracks:
                if track.track_type == 'Video':
                    tmp += [track.width, track.height, track.frame_rate, track.codec, track.encoded_library_name]
                elif track.track_type == 'Audio':
                    tmp += [track.format, track.channel_s, track.bit_rate, track.sampling_rate, track.language]
            ws.append(tmp)

    wb.save(path / f'{args.name.strip()}.xlsx')
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(e)
        sys.exit(1)
