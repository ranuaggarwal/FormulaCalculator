import * as jsonPath from "jsonpath";

export class Utils {
  static documentjson: any;

  public static getDocumentJson(): any {
    return this.documentjson;
  }

  static setDocumentJson(json): any {
    this.documentjson = json;
  }

  static GetInputCells(): any {
    var json = Utils.getDocumentJson();
    var InputCellArray = {};
    var nameRefPathForQuery = "$..['#sheets'][?(@.locked==true)]";
    var sheets = jsonPath.nodes(json, nameRefPathForQuery);
    if (sheets.length > 0) {
      for (let j = 0; j < sheets.length; j++) {
        var node = sheets[j].value;
        var sheetName = node.name;
        nameRefPathForQuery = "$..['#cells'][?(!@.formula)]";
        var cells = jsonPath.nodes(sheets[j], nameRefPathForQuery);
        cells.forEach(element => {
          var eleNode = element.value;
          var cellCompleteName = sheetName + "!" + eleNode.ref;

          InputCellArray[cellCompleteName] =
            eleNode.value == undefined ? 0 : eleNode.value;
        });
      }
      // var sheetName =
      // var InputCellArray = {};
    }
    return InputCellArray;
  }

  static GetOutPutCells(): any {
    var json = Utils.getDocumentJson();
    var OutputCellArray = [];
    var nameRefPathForQuery = "$..['#sheets'][*]";
    var sheets = jsonPath.nodes(json, nameRefPathForQuery);
    if (sheets.length > 0) {
      for (let j = 0; j < sheets.length; j++) {
        var node = sheets[j].value;
        var sheetName = node.name;
        nameRefPathForQuery = "$..['#cells'][?(@.formula)]";
        var cells = jsonPath.nodes(sheets[j], nameRefPathForQuery);
        for (let i = 0; i < cells.length; i++) {
          var eleNode = cells[i].value;
          var outCell = {
            cell: eleNode.ref,
            sheet: sheetName,
            formula: eleNode.formula,
            value: ""
          };

          OutputCellArray.push(outCell);
        }

        //OutputCellArray[cellCompleteName] = eleNode.formula;
      }
    }
    return OutputCellArray;
    // var sheetName =
    // var InputCellArray = {};
  }

  public static GetFormulaIfFormulaCell(cellCoord, sheetName): any {
    var json = Utils.getDocumentJson();
    var elementPath =
      "$['#sheets'][?(@.name=='" +
      sheetName +
      "')]['#rows'][*]['#cells'][?(@.ref == '" +
      cellCoord +
      "')]";
    var variableValue = jsonPath.nodes(json, elementPath);
    if (variableValue.length > 0) {
      if (variableValue[0].value.formula != null) {
        return variableValue[0].value.formula;
      } else {
        return null;
      }
    } else {
      return null;
    }
  }

  static UpdateCellValue(sheet: any, cell: any, result: any) {
    var json = Utils.getDocumentJson();
    var elementPath =
      "$['#sheets'][?(@.name=='" +
      sheet +
      "')]['#rows'][*]['#cells'][?(@.ref == '" +
      cell +
      "')]";
    var variableValue = jsonPath.nodes(json, elementPath);
    if (variableValue.length > 0) {
      variableValue[0].value.value = result;
    }
  }

  public static GetValueOfNodeFromDocumentn(
    row: number,
    col: number,
    sheet: string
  ): any {
    var value = 0;

    var solutionjson = Utils.getDocumentJson();

    //Finding Node and Path for this Submission Cell in SUbmission Json
    var cellRefPathForQuery;

    cellRefPathForQuery =
      "$..['#sheets'][?(@.name=='" +
      sheet +
      "')]['#rows'][" +
      row +
      "]['#cells'][" +
      col +
      "]";

    var Node = jsonPath.nodes(solutionjson, cellRefPathForQuery);

    if (Node.length > 0) {
      var valNode = Node[0].value.value;
    }

    return valNode;
  }

  public static GetCol(index): any {
    var arrChar = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R",
      "S",
      "T",
      "U",
      "V",
      "W",
      "X",
      "Y",
      "Z"
    ];
    return arrChar[index];
  }
}
