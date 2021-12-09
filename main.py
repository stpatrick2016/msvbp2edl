"""
Exports Microsoft Video Editor project into EDL format.
MSVE projects are stored in SQLite database found in this path:
%LocalAppData%\Packages\Microsoft.Windows.Photos_8wekyb3d8bbwe\LocalState\MediaDb.v1.sqlite

More about EDL format here: https://www.niwa.nu/2013/05/how-to-read-an-edl/

Usage:
Scroll to the very bottom and replace the project name you would like to export and output file path
then just run it.

Limitations: only supports cut-type clips (no transitions or special effects)
"""
import json
import os
import sqlite3

import typing
from datetime import timedelta
from math import modf

FRAME_RATE = 30  # assuming our frame rates are 30 fps

DB_FILE_PATH = os.path.expandvars(
    r"%LocalAppData%\Packages\Microsoft.Windows.Photos_8wekyb3d8bbwe\LocalState\MediaDb.v1.sqlite"
)


def _nano100_to_time(nanos: int) -> str:
    """
    Convert MSVE time to EDL time. MSVE stores times in 100 nanoseconds and EDL format
    is hh:mm:ss:ff (ff for frames)
    :param nanos: nanoseconds time to convert
    :return: string in EDL representation
    """
    ns = 10000000  # it is stored in 100ns
    ret = str(timedelta(seconds=nanos // ns))
    if len(ret) < 8:
        ret = f"0{ret}"  # append extra 0 if needed

    frame = int(FRAME_RATE * modf(nanos / ns)[0])
    return f"{ret}:{frame:02}"


def _convert(project_name: str, project_data: typing.Dict, target: str) -> None:
    """
    Given MSVE data, converts and saves into target file
    :param project_name: Name of the project
    :param project_data: Raw data in MSVE format
    :param target: Target file path
    """
    lines = [f"TITLE: {project_name}", "FCM: NON-DROP FRAME", ""]

    index = 1
    master_time = FRAME_RATE

    for card in project_data["Project"]["Cards"]:
        source_path = card["Sources"][0]["MediaBackedSourceProperties"][
            "url"
        ]  # full path to source file
        start_time_raw = card["Sources"][0]["VideoSourceProperties"][
            "idealAssetStartTime"
        ]  # start time in 100 ns
        duration_raw = card["idealDuration"]  # duration in 100 ns

        # format is like this:
        # 001  AX       V     C        00:24:05:25 00:28:34:15 00:00:00:01 00:04:28:20
        times = f"{_nano100_to_time(start_time_raw)} {_nano100_to_time(start_time_raw + duration_raw)} {_nano100_to_time(master_time)} {_nano100_to_time(master_time + duration_raw)}"
        lines.append(f"{index:03}  AX  V  C  {times}")  # video track
        lines.append(f"{index:03}  AX  A  C  {times}")  # audio track

        lines.append(f"* FROM CLIP NAME: {source_path}")
        lines.append("")

        index += 1
        master_time += duration_raw

    with open(target, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def __collate_nocase(s1: str, s2: str):
    """
    MSVE uses custom collation, so we provide it something similar - case-insensitive collation
    """
    if s1.lower() == s2.lower():
        return 0
    elif s1.lower() < s2.lower():
        return 1
    else:
        return -1


def select_project() -> str:
    with sqlite3.connect(f"file:{DB_FILE_PATH}?mode=ro", uri=True) as con:
        con.create_collation("NoCaseLinguistic", __collate_nocase)
        sql = """
        select
            a.Album_Name
        from Album a
        order by a.Album_Name
        """
        cur = con.cursor()
        cur.execute(sql)
        albums = cur.fetchall()
        cur.close()
        while True:
            i = 1
            print("Select which album to convert:")
            for album in albums:
                print(f"\t[{i}]: {album[0]}")
                i += 1

            selection = int(input(">> "))
            if selection and selection in range(1, len(albums) + 1):
                break

        return albums[selection - 1][0]


def select_location() -> str:
    default_location = os.path.join(os.getcwd(), "output.edl")
    location = input(f"Where to save the output (default is {default_location}): ")
    return location or default_location


def generate(project_name: str, target: str):
    with sqlite3.connect(f"file:{DB_FILE_PATH}?mode=ro", uri=True) as con:
        con.create_collation("NoCaseLinguistic", __collate_nocase)

        sql = """
    select 
        p.Project_RpmState
    from Project p
        inner join Album a on p.Project_AlbumId=a.Album_Id
    where a.Album_Name=?
    """
        cur = con.cursor()
        cur.execute(sql, [project_name])

        project = cur.fetchone()
        if project is None:
            raise Exception(f"Project not found in database: {project_name}")

        project_data = json.loads(
            json.loads(project[0])["RenderableProjectManagerBlob"]
        )

    _convert(project_name, project_data, target)


if __name__ == "__main__":
    album = select_project()
    location = select_location()
    generate(album, location)
