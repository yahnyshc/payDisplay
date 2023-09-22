import pathlib
import send2trash
import os
import shutil
import ntpath

class dirclean:
    def __init__(self, testing, log, logsdir):
        self.testing = testing
        self.log = log
        self.logsdir = logsdir

    def setTesting(self, testing):
        self.testing = testing

    def pure_posix_path(self, p):
        return pathlib.PureWindowsPath(p).as_posix()

    def pure_windows_path(self, p):
        return pathlib.PureWindowsPath(p)

    def completely_remove(self, p):
        if not self.testing:
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p)
                elif os.path.isfile(p):
                    os.remove(p)
                else:
                    self.log.write("\nError::unknown file type: " + p + "\n")
            except:
                self.log.write("\nError::problem when deleting: " + p + "\n")

    def to_recycle_bin(self, p):
        if not self.testing:
            try:
                send2trash.send2trash(self.pure_windows_path(p))
            except:
                self.log.write("\nError::problem when deleting: " + p + "\n")

    def move(self, src_file, dest_folder):
        filename = ntpath.basename(src_file)
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        if not self.testing:
            # copy and then move to recycle bin
            shutil.copy(src_file, os.path.join(dest_folder, filename))
            self.to_recycle_bin(src_file)

    def remove_outdated_logs(self, max_logs_amount = 8):
        list_of_files = os.listdir(self.logsdir)

        printed_something = False
        while len(list_of_files) > max_logs_amount:
            full_path = [self.logsdir + r'/{0}'.format(x) for x in list_of_files]
            oldest_file = min(full_path, key=os.path.getctime)
            oldest_file_name = ntpath.basename(oldest_file)

            self.log.write(f"Removing outdated log file => {oldest_file_name}\n")
            printed_something = True

            self.completely_remove(oldest_file)
            list_of_files.remove(oldest_file_name)

        if printed_something:
            self.log.write("\n")