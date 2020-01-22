"""
This is a python library to manage the tables on Google Document using Google Docs API.

- All values can be retrieved from the table on Google Document.
- Values can be put to the table.
- Delete table, rows and columns of the table.
- New table can be created by including values.
- Append rows to the table by including values.
"""

__author__ = "Kanshi TANAIKE (tanaike@hotmail.com)"
__copyright__ = "Copyright 2019, Kanshi TANAIKE"
__license__ = "MIT"
__version__ = "1.1.0"

import os
import sys
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

VERSION = "1.1.0"


def GetTables(resource):
    return gdoctableapp(resource).getTables()


def GetValues(resource):
    return gdoctableapp(resource).getValues()


def DeleteTable(resource):
    return gdoctableapp(resource).deleteTable()


def DeleteRowsAndColumns(resource):
    return gdoctableapp(resource).deleteRowsAndColumns()


def SetValues(resource):
    return gdoctableapp(resource).setValues()


def CreateTable(resource):
    return gdoctableapp(resource).createTable()


def AppendRow(resource):
    return gdoctableapp(resource).appendRow()


def ReplaceTextsToImages(resource):
    return gdoctableapp(resource).replaceTextsToImages()


class gdoctableapp():
    """This is a base class of gdoctableapppy."""

    def __init__(self, resource):
        self.obj = {"params": resource, "result": {"libraryVersion": VERSION}}

        if "showAPIResponse" in resource.keys() and resource["showAPIResponse"]:
            self.obj["result"]["responseFromAPIs"] = []

    def __createInsertInlineImageRequest(self, startIndex, url, width, height):
        if type(width) is str or type(height) is str or width <= 0 and height <= 0:
            try:
                raise ValueError(
                    "Error: Please check 'imageWidth' and 'imageHeight'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        req = {
            "insertInlineImage": {
                "uri": url,
                "location": {"index": startIndex}
            }
        }
        if width > 0 and height > 0:
            req["insertInlineImage"]["objectSize"] = {
                "width": {"magnitude": width, "unit": "PT"},
                "height": {"magnitude": height, "unit": "PT"}
            }
        return req

    def __getTextRunContent(self, ar, h):
        if "paragraph" in h:
            for e in h.get("paragraph").get("elements"):
                if "textRun" in e.keys() and self.obj["params"].get("searchText") in e.get("textRun").get("content"):
                    ar.append(e)

    def __getTableContent(self, e):
        ar = []
        for f in e["table"].get("tableRows"):
            for g in f.get("tableCells"):
                for h in g.get("content"):
                    self.__getTextRunContent(ar, h)
        return ar

    def __deleteTempFile(self):
        if "replaceImageFilePath" in self.obj["params"].keys() and "driveSrv" in self.obj.keys() and self.obj["tempFileId"] != "":
            try:
                self.obj["driveSrv"].files().delete(fileId=self.obj["tempFileId"]).execute()
            except HttpError as err:
                print(err)
                sys.exit(1)

    def __replaceTextsToImagesByURL(self):
        if "searchText" not in self.obj["params"].keys() or self.obj["params"]["searchText"] == "":
            try:
                raise ValueError(
                    "Error: Please set 'searchText'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        if "replaceImageURL" not in self.obj["params"].keys() and "replaceImageFilePath" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'replaceImageURL' or 'replaceImageFilePath'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        allContents = self.__getDocument()
        ar = []
        for e in allContents:
            if "table" in e.keys():
                ar.extend(self.__getTableContent(e))
            elif not self.obj["params"].get("tableOnly") and "paragraph" in e.keys():
                self.__getTextRunContent(ar, e)
        ar.reverse()
        searchText = self.obj["params"].get("searchText")
        replacedUrl = self.obj["params"].get("replaceImageURL")
        width = self.obj["params"].get("imageWidth")
        height = self.obj["params"].get("imageHeight")
        requests = []
        for e in ar:
            content = e["textRun"]["content"]
            if content.strip() == searchText:
                offset = len(content) - len(content.strip())
                requests.append({
                    "deleteContentRange": {
                        "range": {"startIndex": e["startIndex"], "endIndex": e["endIndex"] - offset}
                    }
                })
                requests.append(self.__createInsertInlineImageRequest(e["startIndex"], replacedUrl, width, height))
            else:
                start = e["startIndex"] + content.find(searchText)
                requests.append({
                    "deleteContentRange": {
                        "range": {"startIndex": start, "endIndex": start + len(searchText)}
                    }
                })
                requests.append(self.__createInsertInlineImageRequest(start, replacedUrl, width, height))

        if len(requests) > 0:
            self.obj["requestBody"] = requests
            self.__documentsBatchUpdate()
        else:
            self.obj["result"]["message"] = "'" + searchText + "' was not found."

    def __uploadImageFile(self):
        p = self.obj["params"]
        if ("replaceImageURL" not in p.keys() or p["replaceImageURL"] == "") and "replaceImageFilePath" in p.keys() and p["replaceImageFilePath"] is not "":
            c = p["oauth2"] if "oauth2" in p.keys() else p["service_account"] if "service_account" in p.keys() else None
            if not c:
                try:
                    raise ValueError(
                        "Error: You can use API key, OAuth2 and Service account.")
                except ValueError as err:
                    print(err)
                    sys.exit(1)

            try:
                self.obj["driveSrv"] = build("drive", "v3", credentials=c)
                metadata = {"name": os.path.basename(p["replaceImageFilePath"])}
                media = MediaFileUpload(p["replaceImageFilePath"])
                r1 = self.obj["driveSrv"].files().create(body=metadata, media_body=media, fields="id,webContentLink").execute()
                self.obj["params"]["replaceImageURL"] = r1["webContentLink"]
                self.obj["tempFileId"] = r1["id"]
                if "showAPIResponse" in p.keys() and p["showAPIResponse"]:
                    self.obj["result"]["responseFromAPIs"].append(r1)
                permission = {"type": "anyone", "role": "reader"}
                r2 = self.obj["driveSrv"].permissions().create(fileId=r1["id"], body=permission, fields="id").execute()
                if "showAPIResponse" in p.keys() and p["showAPIResponse"]:
                    self.obj["result"]["responseFromAPIs"].append(r2)
            except HttpError as err:
                print(err)
                sys.exit(1)

    def __replaceTextsToImagesMain(self):
        self.__uploadImageFile()
        self.__replaceTextsToImagesByURL()
        self.__deleteTempFile()

    def __appendRowMain(self):
        if "values" not in self.obj["params"].keys() or not self.obj["params"]["values"]:
            try:
                raise ValueError(
                    "Error: Values for putting are not set.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.obj["params"]["values"] = [{
            "values": self.obj["params"]["values"],
            "range": {"startRowIndex": self.obj["docTable"]["table"]["rows"], "startColumnIndex": 0}
        }]
        self.__setValuesMain()

    def __createTableMain(self):
        if "createIndex" not in self.obj["params"].keys() and "append" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'createIndex' or 'append'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        if "rows" not in self.obj["params"].keys() or "columns" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set rows and columns for creating new table.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        if "append" in self.obj["params"].keys():
            self.__appendTable()
        elif "createIndex" in self.obj["params"].keys():
            self.__insertTable()

    def __insertTable(self):
        requests = []
        requests.append({
            "insertTable": {
                "rows": self.obj["params"]["rows"],
                "columns": self.obj["params"]["columns"],
                "location": {
                    "index": self.obj["params"]["createIndex"]
                }
            }
        })

        if "values" in self.obj["params"].keys() and self.obj["params"]["values"]:
            self.__createRequestBodyForInsertText(
                requests, self.obj["params"]["createIndex"])

        self.obj["requestBody"] = requests
        self.__documentsBatchUpdate()

    def __appendTable(self):
        self.obj["requestBody"] = [
            {
                "insertTable": {
                    "rows": self.obj["params"]["rows"],
                    "columns": self.obj["params"]["columns"],
                    "endOfSegmentLocation": {"segmentId": ""}
                }
            }
        ]
        self.__documentsBatchUpdate()
        contents = self.__getDocument()
        for content in reversed(contents):
            if "table" in content.keys():
                self.obj["docTable"] = content
                break

        requests = []
        self.__createRequestBodyForInsertText(
            requests,
            self.obj["docTable"]["startIndex"] - 1
        )
        self.obj["requestBody"] = requests
        self.__documentsBatchUpdate()

    def __createRequestBodyForInsertText(self, requests, idx):
        val = self.__parseInputValues(
            self.obj["params"]["values"],
            idx,
            self.obj["params"]["rows"],
            self.obj["params"]["columns"]
        )
        for e in reversed(val):
            v = e["content"]
            if v != "":
                requests.append({
                    "insertText": {
                        "location": {
                            "index": e["index"]
                        },
                        "text": v
                    }
                })

    def __parseInputValues(self, values, index, rows, cols):
        index += 4
        v = []
        maxRow = len(values)
        for row in range(rows):
            if maxRow > row:
                maxCol = len(values[row])
            else:
                maxCol = cols
            for col in range(cols):
                if maxRow > row and maxCol > col and values[row][col] != "":
                    v.append({
                        "row": row,
                        "col": col,
                        "content": values[row][col],
                        "index": index
                    })
                index += 2
            index += 1
        return v

    def __setValuesMain(self):
        if "values" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'values'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        if self.__valuesChecker():
            self.obj["params"]["values"] = [
                {
                    "values": self.obj["params"]["values"],
                    "range": {"startRowIndex": 0, "startColumnIndex": 0}
                }
            ]

        dupChk = self.__checkDupValues()
        if dupChk["dup"]:
            try:
                raise ValueError(
                    "Error: Range of inputted values are duplicated.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.obj["requests"] = {}
        self.__parseInputValuesForSetValues(dupChk)
        self.__addRowsAndColumnsForSetValues()
        self.__addRowsAndColumnsByAPI()
        self.__parseTable()
        self.obj["requestBody"] = self.__createRequestsForSetValues()
        self.__documentsBatchUpdate()

    def __deleteRowsAndColumnsMain(self):
        maxDeleteRows = max(self.obj["params"]["deleteRows"]) + 1
        maxDeleteCols = max(self.obj["params"]["deleteColumns"]) + 1
        table = self.obj["docTable"]["table"]
        if table["rows"] < maxDeleteRows or table["columns"] < maxDeleteCols:
            try:
                raise ValueError(
                    "Error: Rows and columns for deleting are outside of the table.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        tablePos = self.obj["docTable"]["startIndex"]
        iObj = self.obj["params"]
        requests = []
        if "deleteRows" in iObj.keys() and isinstance(iObj["deleteRows"], list) and iObj["deleteRows"]:
            iObj["deleteRows"].sort(reverse=True)
            for e in iObj["deleteRows"]:
                requests.append({"deleteTableRow": {"tableCellLocation": {"tableStartLocation": {"index": tablePos}, "rowIndex": e}}})

        if "deleteColumns" in iObj.keys() and isinstance(iObj["deleteColumns"], list) and iObj["deleteColumns"]:
            iObj["deleteColumns"].sort(reverse=True)
            for e in iObj["deleteColumns"]:
                requests.append({"deleteTableColumn": {"tableCellLocation": {"tableStartLocation": {"index": tablePos}, "columnIndex": e}}})

        self.obj["requestBody"] = requests
        self.__documentsBatchUpdate()

    def __deleteTableMain(self):
        self.obj["requestBody"] = [
            {
                "deleteContentRange": {
                    "range": {
                        "startIndex": self.obj["docTable"]["startIndex"],
                        "endIndex": self.obj["docTable"]["endIndex"]
                    }
                }
            }
        ]
        self.__documentsBatchUpdate()

    def __getValuesFromTable(self):
        self.__parseTable()
        content = self.obj["content"]
        values = []
        for e in content:
            row = e
            temp1 = []
            for j, f in enumerate(row):
                temp2 = []
                for k, _ in enumerate(f):
                    temp2.append(row[j][k]["content"].replace("\n", ""))
                temp1.append("\n".join(temp2))
            values.append(temp1)
        return values

    def __getValuesMain(self):
        values = self.__getValuesFromTable()
        self.obj["result"]["values"] = values

    def __getTablesMain(self):
        tables = []
        for i, table in enumerate(self.obj["docTables"]):
            self.obj["docTable"] = table
            values = self.__getValuesFromTable()
            tables.append({
                "index": i,
                "values": values,
                "tablePosition": {
                    "startIndex": table["startIndex"],
                    "endIndex": table["endIndex"]
                }
            })

        self.obj["result"]["tables"] = tables

    def __valuesChecker(self):
        return all([isinstance(e, list) for e in self.obj["params"].get("values")])

    def __checkDupValues(self):
        values = self.obj["params"]["values"]
        temp = []
        for e in values:
            rowOffset = e["range"]["startRowIndex"]
            colOffset = e["range"]["startColumnIndex"]
            temp1 = []
            for i, row in enumerate(e["values"]):
                temp2 = []
                for j, col in enumerate(row):
                    temp2.append(
                        {"row": i + rowOffset, "col": j + colOffset, "content": col})
                temp1.extend(temp2)
            temp.extend(temp1)
        dupCheck = {"dup": [], "noDup": []}
        for e in temp:
            if any(True if f["row"] == e["row"] and f["col"] == e["col"] else False for f in dupCheck["noDup"]):
                dupCheck["dup"].append(e)
            else:
                dupCheck["noDup"].append(e)
        return dupCheck

    def __parseInputValuesForSetValues(self, dupChk):
        s1 = sorted(dupChk["noDup"], key=lambda x: x["col"])
        s2 = sorted(s1, key=lambda x: x["row"])
        self.obj["parsedValues"] = s2

    def __addRowsAndColumns(self, startIndex, maxRow, maxCol, tableRow, tableCol):
        addRows = maxRow - tableRow
        addColumns = maxCol - tableCol
        obj = {"insertTableRowBody": [], "insertTableColumnBody": []}
        if addRows > 0:
            for i in range(addRows):
                obj["insertTableRowBody"].append({
                    "insertTableRow": {
                        "insertBelow": True,
                        "tableCellLocation": {
                            "tableStartLocation": {"index": startIndex},
                            "rowIndex": tableRow - 1 + i
                        }
                    }
                })

        if addColumns > 0:
            for i in range(addColumns):
                obj["insertTableColumnBody"].append({
                    "insertTableColumn": {
                        "insertRight": True,
                        "tableCellLocation": {
                            "tableStartLocation": {"index": startIndex},
                            "columnIndex": tableCol - 1 + i
                        }
                    }
                })

        return obj

    def __addRowsAndColumnsForSetValues(self):
        values = self.obj["params"]["values"]
        res = {"maxRow": 0, "maxCol": 0}
        for e in values:
            maxRow = len(e["values"]) + e["range"]["startRowIndex"]
            maxCol = 0
            for f in e["values"]:
                l = len(f)
                if maxCol < l:
                    maxCol = l
            maxCol += e["range"]["startColumnIndex"]
            if res["maxRow"] < maxRow:
                res["maxRow"] = maxRow
            if res["maxCol"] < maxCol:
                res["maxCol"] = maxCol

        o = self.__addRowsAndColumns(
            self.obj["docTable"]["startIndex"],
            res["maxRow"],
            res["maxCol"],
            self.obj["docTable"]["table"]["rows"],
            self.obj["docTable"]["table"]["columns"]
        )
        self.obj["requests"]["insertTableRow"] = o["insertTableRowBody"]
        self.obj["requests"]["insertTableColumn"] = o["insertTableColumnBody"]

    def __addRowsAndColumnsByAPI(self):
        tr = self.obj["requests"]["insertTableRow"]
        tc = self.obj["requests"]["insertTableColumn"]
        if tr or tc:
            requests = []
            if tr:
                for e in tr:
                    requests.append(e)
            if tc:
                for e in tc:
                    requests.append(e)
            self.obj["requestBody"] = requests
            self.__documentsBatchUpdate()
            self.__getTable()

    def __createRequestsForSetValues(self):
        requests = []
        values = self.obj["parsedValues"]
        for e in reversed(values):
            r = e["row"]
            c = e["col"]
            v = str(e["content"])
            delReq = self.obj["delCell"][r][c]
            if delReq["deleteContentRange"]["range"]["startIndex"] != delReq["deleteContentRange"]["range"]["endIndex"]:
                requests.append(delReq)
            if v != "":
                requests.append({
                    "insertText": {
                        "location": {"index": delReq["deleteContentRange"]["range"]["startIndex"]},
                        "text": v
                    }
                })
        return requests

    def __parseTable(self):
        docContent = self.obj["docTable"]
        tableRows = docContent["table"].get("tableRows")
        valuesIndexes = {"deleteIndex": [], "content": []}
        for e in tableRows:
            tableCells = e.get("tableCells")
            tempRowsDelCell = []
            tempRowsContent = []
            for f in tableCells:
                tempColsDelCell = {"deleteContentRange": {"range": {}}}
                tempColsContent = []
                contents = f["content"]
                for k, g in enumerate(contents):
                    if "paragraph" in g.keys():
                        elements = g["paragraph"]["elements"]
                        for l, h in enumerate(elements):
                            if k == 0 and l == 0:
                                tempColsDelCell["deleteContentRange"]["range"]["startIndex"] = h["startIndex"]
                            if k == len(contents) - 1 and l == len(elements) - 1:
                                tempColsDelCell["deleteContentRange"]["range"]["endIndex"] = h["endIndex"] - 1
                            cellContent = ""
                            if "textRun" in h.keys():
                                cellContent = h["textRun"]["content"]
                            elif "inlineObjectElement" in h.keys():
                                cellContent = "[INLINE OBJECT]"
                            else:
                                cellContent = "[UNSUPPORTED CONTENT]"
                            tempColsContent.append({
                                "startIndex": h["startIndex"],
                                "endIndex": h["endIndex"],
                                "content": cellContent
                            })
                    elif "table" in g.keys():
                        tempColsContent.append({
                            "startIndex": g["startIndex"],
                            "endIndex": g["endIndex"],
                            "content": "[TABLE]"
                        })
                    else:
                        tempColsContent.append({
                            "startIndex": g["startIndex"],
                            "endIndex": g["endIndex"],
                            "content": "[UNSUPPORTED CONTENT]"
                        })

                tempRowsDelCell.append(tempColsDelCell)
                tempRowsContent.append(tempColsContent)

            valuesIndexes["deleteIndex"].append(tempRowsDelCell)
            valuesIndexes["content"].append(tempRowsContent)

        self.obj["delCell"] = valuesIndexes["deleteIndex"]
        self.obj["content"] = valuesIndexes["content"]
        self.obj["cell1stIndex"] = valuesIndexes["content"][0][0][0]["startIndex"]

    def __documentsBatchUpdate(self):
        try:
            r = self.obj["service"].documents().batchUpdate(
                documentId=self.obj["params"]["documentId"],
                body={"requests": self.obj["requestBody"]}).execute()
            if "showAPIResponse" in self.obj["params"].keys() and self.obj["params"]["showAPIResponse"]:
                self.obj["result"]["responseFromAPIs"].append(r)
            self.obj["requestBody"] = []
        except HttpError as err:
            print(err)
            sys.exit(1)

    def __getDocument(self):
        try:
            doc = self.obj["service"].documents().get(
                documentId=self.obj["params"]["documentId"]).execute()
            contents = doc.get("body").get("content")
            if "showAPIResponse" in self.obj["params"].keys() and self.obj["params"]["showAPIResponse"]:
                self.obj["result"]["responseFromAPIs"].append(contents)
            return contents
        except HttpError as err:
            print(err)
            sys.exit(1)

    def __getAllTables(self):
        contents = self.__getDocument()
        tables = []
        for content in contents:
            if "table" in content.keys():
                tables.append(content)
        self.obj["docTables"] = tables

    def __getTable(self):
        contents = self.__getDocument()
        ti = 0
        table = {}
        for content in contents:
            if "table" in content.keys():
                if ti == self.obj["params"].get("tableIndex"):
                    table = content
                    break
                ti += 1

        if not table.keys():
            try:
                raise ValueError("Error: No table is found at index " +
                                 str(self.obj["params"].get("tableIndex")))
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.obj["docTable"] = table

    def __getService(self):
        if "oauth2" in self.obj["params"].keys():
            return build('docs', 'v1', credentials=self.obj["params"]["oauth2"])
        if "service_account" in self.obj["params"].keys():
            return build('docs', 'v1', credentials=self.obj["params"]["service_account"])
        try:
            raise ValueError(
                "Error: You can use API key, OAuth2 and Service account.")
        except ValueError as err:
            print(err)
            sys.exit(1)

    def __init(self):
        if self.obj["params"].get("documentId") == "":
            try:
                raise ValueError(
                    "Error: Please set 'documentId'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.obj["service"] = self.__getService()
        if "tableIndex" in self.obj["params"].keys():
            if self.obj["params"]["tableIndex"] == -1:
                self.__getAllTables()
            else:
                self.__getTable()

    def getTables(self):
        self.obj["params"]["tableIndex"] = -1
        self.__init()
        self.__getTablesMain()
        return self.obj["result"]

    def getValues(self):
        if "tableIndex" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'tableIndex'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.__init()
        self.__getValuesMain()
        return self.obj["result"]

    def deleteTable(self):
        if "tableIndex" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'tableIndex'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.__init()
        self.__deleteTableMain()
        return self.obj["result"]

    def deleteRowsAndColumns(self):
        if "tableIndex" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'tableIndex'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.__init()
        self.__deleteRowsAndColumnsMain()
        return self.obj["result"]

    def setValues(self):
        if "tableIndex" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'tableIndex'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.__init()
        self.__setValuesMain()
        return self.obj["result"]

    def createTable(self):
        self.__init()
        self.__createTableMain()
        return self.obj["result"]

    def appendRow(self):
        if "tableIndex" not in self.obj["params"].keys():
            try:
                raise ValueError(
                    "Error: Please set 'tableIndex'.")
            except ValueError as err:
                print(err)
                sys.exit(1)

        self.__init()
        self.__appendRowMain()
        return self.obj["result"]

    def replaceTextsToImages(self):
        self.__init()
        self.__replaceTextsToImagesMain()
        return self.obj["result"]
