import * as jsonPath from 'jsonpath';

export class Utils {
  static documentjson: any;

  public static getDocumentJson(): any {
    return this.documentjson;
  }

  static setDocumentJson(json): any {
    this.documentjson = json;
  }

  static GetInputCells(): any {
    const json = Utils.getDocumentJson();
    let InputCellArray = {};
    let nameRefPathForQuery = "$..['#sheets'][?(@.locked==true)]";
    let sheets = jsonPath.nodes(json, nameRefPathForQuery);
    if (sheets.length > 0) {
      for (let j = 0; j < sheets.length; j++) {
        let node = sheets[j].value;
        let sheetName = node.name;
        nameRefPathForQuery = "$..['#cells'][?(!@.formula)]";
        var cells = jsonPath.nodes(sheets[j], nameRefPathForQuery);
        cells.forEach(element => {
          let eleNode = element.value;
          let cellCompleteName = sheetName + '!' + eleNode.ref;

          InputCellArray[cellCompleteName] = eleNode.value == undefined ? 0 : eleNode.value;
        });
      }
      // var sheetName =
      // var InputCellArray = {};
    }
    return InputCellArray;
  }

  static GetOutPutCells(): any {
    const json = Utils.getDocumentJson();
    let OutputCellArray = [];
    let nameRefPathForQuery = "$..['#sheets'][*]";
    let sheets = jsonPath.nodes(json, nameRefPathForQuery);
    if (sheets.length > 0) {
      for (let j = 0; j < sheets.length; j++) {
        let node = sheets[j].value;
        let sheetName = node.name;
        nameRefPathForQuery = "$..['#cells'][?(@.formula)]";
        let cells = jsonPath.nodes(sheets[j], nameRefPathForQuery);
        for (let i = 0; i < cells.length; i++) {
          let eleNode = cells[i].value;
          let outCell = {
            cell: eleNode.ref,
            sheet: sheetName,
            formula: eleNode.formula,
            value: '',
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
    const json = Utils.getDocumentJson();
    let elementPath =
      "$['#sheets'][?(@.name=='" + sheetName + "')]['#rows'][*]['#cells'][?(@.ref == '" + cellCoord + "')]";
    let variableValue = jsonPath.nodes(json, elementPath);
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
    const json = Utils.getDocumentJson();
    var elementPath = "$['#sheets'][?(@.name=='" + sheet + "')]['#rows'][*]['#cells'][?(@.ref == '" + cell + "')]";
    var variableValue = jsonPath.nodes(json, elementPath);
    if (variableValue.length > 0) {
      variableValue[0].value.value = result;
    }
  }

  public static GetValueOfNodeFromDocumentn(row: number, col: number, sheet: string): any {
    let valNode = 0;

    let solutionjson = Utils.getDocumentJson();

    //Finding Node and Path for this Submission Cell in SUbmission Json
    let cellRefPathForQuery;

    cellRefPathForQuery = "$..['#sheets'][?(@.name=='" + sheet + "')]['#rows'][" + row + "]['#cells'][" + col + ']';

    let Node = jsonPath.nodes(solutionjson, cellRefPathForQuery);

    if (Node.length > 0) {
      valNode = Node[0].value.value;
    }

    return valNode;
  }

  public static GetCol(index): any {
    let arrChar = [
      'A',
      'B',
      'C',
      'D',
      'E',
      'F',
      'G',
      'H',
      'I',
      'J',
      'K',
      'L',
      'M',
      'N',
      'O',
      'P',
      'Q',
      'R',
      'S',
      'T',
      'U',
      'V',
      'W',
      'X',
      'Y',
      'Z',
    ];
    return arrChar[index];
  }
}
