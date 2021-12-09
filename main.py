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
import os
import typing

from ms_photos import MsPhotosDb


def select_location() -> str:
    default_location = os.path.join(os.getcwd(), "output.edl")
    location = input(f"Where to save the output (default is {default_location}): ")
    return location or default_location


def select_project(albums: typing.List[str]) -> str:
    while True:
        i = 1
        print("Select which album to convert:")
        for album in albums:
            print(f"\t[{i}]: {album}")
            i += 1

        selection = int(input(">> "))
        if selection and selection in range(1, len(albums) + 1):
            break

    return albums[selection - 1]


if __name__ == "__main__":
    db = MsPhotosDb()
    album = select_project(db.get_projects())
    location = select_location()
    db.export_edl(album, location)
