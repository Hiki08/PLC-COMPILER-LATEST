#%%
from Imports import *

class filesReader():
    fileFinishedReading = ""
    readingYear = ""

    #Quality Control Data
    EM0580106PData = []

    def __init__(self):
        pass
    def ReadEm2pFiles(self):
        while not self.fileFinishedReading:
            try:
                vt1Directory = (fr'\\192.168.2.19\quality control\{str(self.readingYear)}')

                for d in os.listdir(vt1Directory):
                    if "supplier" in d.lower():
                        vt1Directory = os.path.join(vt1Directory, d)
                        for d in os.listdir(vt1Directory):
                            if "inspection standard" in d.lower():
                                vt1Directory = os.path.join(vt1Directory, d)
                                for d in os.listdir(vt1Directory):
                                    if "receiving inspection record" in d.lower():
                                        vt1Directory = os.path.join(vt1Directory, d)

                                        #GETTING EM0580106P FILES
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "gaptec" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    #Finding A Folder That Contains New Trend
                                                    for d in os.listdir(directory):
                                                        if 'new trend' in d.lower():
                                                            directory = os.path.join(directory, d)
                                                            print(f"Updated vt1Directory: {directory}")
                                                            break

                                                    os.chdir(directory)

                                                    files = glob.glob('*EM0580106P*.xlsm')
                                                    print(files)

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        self.EM0580106PData.append(em2PData)
                                        except:
                                            print("NO DATA FOUND IN GAPTEC")
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "dhye" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    #Finding A Folder That Contains New Trend
                                                    for d in os.listdir(directory):
                                                        if 'new trend' in d.lower():
                                                            directory = os.path.join(directory, d)
                                                            print(f"Updated vt1Directory: {directory}")
                                                            break

                                                    os.chdir(directory)

                                                    files = glob.glob('*EM0580106P*.xlsm')
                                                    print(files)

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        self.EM0580106PData.append(em2PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")

            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                self.fileFinishedReading = True

#%%
filesreader = filesReader()
filesreader.readingYear = 2025
filesreader.ReadEm2pFiles()

len(filesreader.EM0580106PData)

# %%
