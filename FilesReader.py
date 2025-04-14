#%%
from Imports import *
import DateAndTimeManager

#EM2P
EM0580106PData = []
EM0660046PData = []
EM0660044PData = []

#EM3P
EM0580107PData = []
EM0660047PData = []
EM0660045PData = []

#FM
FM05000102Data = []

#CSB
CSB6400802Data = []

class filesReader():
    global EM0580106PData
    global EM0660046PData
    global EM0660044PData

    global EM0580107PData
    global EM0660047PData
    global EM0660045PData 

    #RESETING VALUES
    EM0580106PData = []
    EM0660046PData = []
    EM0660044PData = []

    EM0580107PData = []
    EM0660047PData = []
    EM0660045PData = []

    readingYearStored = ""
    readingYear = ""

    def __init__(self):
        pass
    def ReadEm2pFiles(self):
        self.readingYear = self.readingYearStored

        while True:
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
                                                        EM0580106PData.append(em2PData)
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
                                                        EM0580106PData.append(em2PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")






                                        #GETTING EM0660046P FILES
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

                                                    files = glob.glob('*EM0660046P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660046PData.append(em2PData)
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

                                                    files = glob.glob('*EM0660046P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660046PData.append(em2PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")






                                        #GETTING EM0660044P FILES
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

                                                    files = glob.glob('*EM0660044P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660044PData.append(em2PData)
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

                                                    files = glob.glob('*EM0660046P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em2PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em2PData = pd.DataFrame(em2PData)
                                                        em2PData = em2PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM2P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660044PData.append(em2PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")
                                        

            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                # self.fileFinishedReading = True
                break

        #REPLACING BLANK VALUES WITH N/A
        for file in EM0580106PData:
            file.replace('', np.nan, inplace=True)  
        for file in EM0660046PData:
            file.replace('', np.nan, inplace=True)  
        for file in EM0660044PData:
            file.replace('', np.nan, inplace=True)  

    def ReadEm3pFiles(self):
        self.readingYear = self.readingYearStored

        while True:
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

                                        #GETTING EM0580107P FILES
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

                                                    files = glob.glob('*EM0580107P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0580107PData.append(em3PData)
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

                                                    files = glob.glob('*EM0580107P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0580107PData.append(em3PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")






                                        #GETTING EM0660047P FILES
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

                                                    files = glob.glob('*EM0660047P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660047PData.append(em3PData)
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

                                                    files = glob.glob('*EM0660047P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660047PData.append(em3PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")




                                        #GETTING EM0660045P FILES
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

                                                    files = glob.glob('*EM0660045P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660045PData.append(em3PData)
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

                                                    files = glob.glob('*EM0660045P*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        em3PData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        em3PData = pd.DataFrame(em3PData)
                                                        em3PData = em3PData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"EM3P FINDED IN {self.readingYear} NEW TREND")
                                                        EM0660045PData.append(em3PData)
                                        except:
                                            print("NO DATA FOUND IN DHYE")


            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                # self.fileFinishedReading = True
                break

        #REPLACING BLANK VALUES WITH N/A
        for file in EM0580107PData:
            file.replace('', np.nan, inplace=True)
        for file in EM0660047PData:
            file.replace('', np.nan, inplace=True)  
        for file in EM0660045PData:
            file.replace('', np.nan, inplace=True)  

    def ReadFmFiles(self):
        self.readingYear = self.readingYearStored

        while True:
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



                                        # GETTING FM05000102 FILES
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "cronics" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    #Finding A Folder That Contains New Trend
                                                    for d in os.listdir(directory):
                                                        if 'new trend' in d.lower():
                                                            directory = os.path.join(directory, d)
                                                            print(f"Updated vt1Directory: {directory}")
                                                            break

                                                    os.chdir(directory)

                                                    files = glob.glob('*FM05000102*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        fmData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        fmData = pd.DataFrame(fmData)
                                                        fmData = fmData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"FM FINDED IN {self.readingYear} NEW TREND")
                                                        FM05000102Data.append(fmData)
                                        except:
                                            print("NO DATA FOUND IN CRONICS")
            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                # self.fileFinishedReading = True
                break

        #REPLACING BLANK VALUES WITH N/A
        for file in FM05000102Data:
            file.replace('', np.nan, inplace=True)

    def ReadDfbFiles(self):
        pass















    def ReadCsbFiles(self):
        self.readingYear = self.readingYearStored

        while True:
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

                                        #GETTING CSB6400802 FILES
                                        try:
                                            for d in os.listdir(vt1Directory):
                                                if "cronics" in d.lower():
                                                    directory = os.path.join(vt1Directory, d)

                                                    #Finding A Folder That Contains New Trend
                                                    for d in os.listdir(directory):
                                                        if 'new trend' in d.lower():
                                                            directory = os.path.join(directory, d)
                                                            print(f"Updated vt1Directory: {directory}")
                                                            break

                                                    os.chdir(directory)

                                                    files = glob.glob('*CSB6400802*.xlsm')

                                                    for f in files:
                                                        print(f'File Readed {f}')
                                                        workbook = CalamineWorkbook.from_path(f)

                                                        csbData = workbook.get_sheet_by_name("format").to_python(skip_empty_area=True)
                                                        csbData = pd.DataFrame(csbData)
                                                        csbData = csbData.replace(r'\s+', '', regex=True)
                                                        
                                                        print(f"CSB FINDED IN {self.readingYear} NEW TREND")
                                                        CSB6400802Data.append(csbData)
                                        except:
                                            print("NO DATA FOUND IN CRONICS")

            except:
                pass

            if self.readingYear > 2021:
                self.readingYear -= 1
            else:
                # self.fileFinishedReading = True
                break

        #REPLACING BLANK VALUES WITH N/A
        for file in CSB6400802Data:
            file.replace('', np.nan, inplace=True)

#%%
# filesreader = filesReader()
# filesreader.readingYear = 2025
# filesreader.ReadEm2pFiles()

# print(len(EM0580106PData))
# print(len(EM0660046PData))
# print(len(EM0660044PData))

# %%
# filesreader = filesReader()
# filesreader.readingYearStored = 2025
# filesreader.ReadEm3pFiles()

# print(len(EM0580107PData))
# print(len(EM0660047PData))
# print(len(EM0660045PData))

#%%
# filesreader = filesReader()
# filesreader.readingYearStored = 2025
# filesreader.ReadFmFiles()

# print(len(FM05000102Data))
# %%
# filesreader = filesReader()
# filesreader.readingYearStored = 2025
# filesreader.ReadCsbFiles()

# print(len(CSB6400802Data))
#%%