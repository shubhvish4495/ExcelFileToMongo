import xlrd
import configuration as cfg   
import mongoConnect as mongo 

connection =  mongo.connect()

def processFile(fileName):
    try:
        wb = xlrd.open_workbook(fileName)
        fileNameCollection = connection[cfg.EXCEL_FILES_PROCESSED_COLLECTION]
        processedFiles = fileNameCollection.find()
        for files in processedFiles:
            if files["file_name"] == fileName :
                print("################# File Already Processed ################# \n Exiting")
                exit()
        print("################### File Processing Started ###################")
        sheet = wb.sheet_by_index(0)
        columns = sheet.row_values(0)
        arrayOfDict=[]
        for colNum in range(1, sheet.nrows):
            dictEle = {}
            colDataList = sheet.row_values(colNum)
            for i in range(len(colDataList)):
                dictEle[columns[i]] = colDataList[i]
            arrayOfDict.append(dictEle)
        collegeDataCollection = connection[cfg.COLLEGE_DATA_COLLECTION]
        collegeDataCollection.insert_many(arrayOfDict)
        print("################# Records inserted in DB Successfully #################")
        fileNameCollection.insert_one({'file_name':fileName})
    except Exception as e:
        print(str(e))


if __name__ == "__main__":
    fileName = "open_close_rank_wb.xlsx"
    processFile(fileName)
