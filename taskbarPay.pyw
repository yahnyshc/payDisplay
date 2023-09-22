import pystray
import PIL.Image
import os
import calendar
import pandas
import datetime
import ntpath
import re
import time
from modules.dirclean import dirclean

script_name = os.path.splitext(ntpath.basename(__file__))[0]
logs_dir = os.path.join(r"D:\payDisplay\logs", script_name)

prev_title = None
def start_log(d, logs_dir, script_name):
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
    log_name = os.path.join(logs_dir, script_name + d.strftime("%Y-%m-%d_%H-%M") + '.txt')
    return open(log_name, 'w', encoding='utf-8')

def sort_work_timetables():
    d = datetime.datetime.now().replace(microsecond=0)

    log = start_log(d=d, logs_dir=logs_dir, script_name=script_name)

    log.write(f'Started timetables sorting on {d}\n\n')

    cleaner = dirclean(testing=False, log=log, logsdir=logs_dir)

    cleaner.remove_outdated_logs(4)

    downloads_folder = r"C:\Users\yagni\Downloads"
    timetabes_folder = r'D:\work_timetables'

    timetables = list(filter(lambda file: re.search("Maksym.*\d\.xlsx$", file), os.listdir(downloads_folder)))
    timetables = [(downloads_folder + r'/{0}'.format(x)) for x in timetables]
    printed_something = False
    for file in timetables:
        printed_something = True
        try:
            day = datetime.datetime.strptime(file[-13:-5], "%Y%m%d")
        except:
            log.write("Error::problems when converting timetable date: " + file + "\n\n")
            return
        if day.weekday() != 0:
            log.write(f'Error::timetable is not starting on monday => {file}\n\n')
        else:
            destination_year = os.path.join(timetabes_folder, str(day.year))
            destination_month = os.path.join(destination_year, calendar.month_name[int(day.month)])

            log.write(f'Moving timetable:\n')
            final_destination = os.path.join(destination_month, ntpath.basename(file))
            log.write(f'{cleaner.pure_posix_path(file)} => {cleaner.pure_posix_path(final_destination)}\n')

            cleaner.move(file, destination_month)

    if printed_something:
        log.write("\n")

    # manage rest of the timetables
    rest_of_timetables = list(filter(lambda file: re.search("Maksym.*\)\.xlsx$", file), os.listdir(downloads_folder)))
    rest_of_timetables = [(downloads_folder + r'/{0}'.format(x)) for x in rest_of_timetables]
    printed_something = False
    for file in rest_of_timetables:
        printed_something = True
        log.write(f"Moving timetable to recycle bin:\n")
        if os.path.isfile(file):
            log.write(cleaner.pure_posix_path(file) + "\n")
            cleaner.to_recycle_bin(cleaner.pure_windows_path(file))
        else:
            log.write(f"Error::timetable is not file: " + file + "\n")

    if printed_something:
        log.write("\n")

    log.write(f'Finished timetables sorting.\n')
    log.close()

def on_clicked(icon, item = None):
    sort_work_timetables()

    timetables_dir = r'D:\work_timetables'

    years = os.listdir(timetables_dir)

    latest_year = max(years, key=int)

    months = os.listdir(os.path.join(timetables_dir, latest_year))

    latest_month = None
    for m in calendar.month_name[::-1]:
        if m in months:
            latest_month = m
            break

    latest_month_timetables = os.path.join(timetables_dir, latest_year, latest_month)

    total = 0

    os.chdir(latest_month_timetables)
    for timetable in os.listdir(latest_month_timetables):
        df = pandas.read_excel(timetable, engine='openpyxl')
        total += float(df.iloc[12,10])

    hourly_rate = 25
    prev_title = icon.title
    icon.title = "£" + str(total*hourly_rate)
    if item is not None:
        if prev_title == icon.title:
            icon.notify("Current", icon.title)
        else:
            #difference = "£" + str(float(icon.title[1:])-float(prev_title[1:]))
            icon.notify(
                "Updated",
                prev_title+" => "+icon.title)
        time.sleep(4.5)
        icon.remove_notification()




if __name__ == "__main__":
    image = PIL.Image.open(r"D:\payDisplay\images\pound_icon.png")

    icon = pystray.Icon("Pay", image, menu=pystray.Menu(
        pystray.MenuItem("Update", on_clicked, default=True, visible=False)
    ))

    on_clicked(icon)

    icon.run()
