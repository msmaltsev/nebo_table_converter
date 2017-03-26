import shutil, os, datetime
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-p",
                    dest="purge",
                    default=False,
                    action='store_true',
                    help="Clear previous versions")
args = parser.parse_args()


if args.purge:
    print('AND THEN THERE WAS A PURGE\nAll previous versions deleted')
    zips = [f for f in os.listdir(os.getcwd()) if os.path.splitext(f)[-1] == '.zip']
    for z in zips:
        print(z)
        os.remove(z)
else:
    print('All previous versions are still available')

now = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
shutil.make_archive('excel_FINAL_%s'%now, 'zip', 'build')
