import * as formulaParser from "libs-leonardo-formula-parser";
import { Utils } from "../Common/Utils";
import { FormulaUtils } from "../Common/FormulaUtils";

export class DefaultFormulaEngine {
  public EvaluateFormula(documentJson: any) {
    Utils.setDocumentJson(documentJson);
    //Get all input cells
    let InputCellColls = Utils.GetInputCells();

    //Get all output cells
    //All formula cells
    let OutputCellColls = Utils.GetOutPutCells();

    OutputCellColls.forEach(element => {
      var result = null;
      if (!InputCellColls[element.sheet + "!" + element.cell]) {
        result = FormulaUtils.EvaluateFormula(
          element.formula,
          element.sheet,
          InputCellColls
        );

        InputCellColls[element.sheet + "!" + element.cell] = result;
      } else {
        result = InputCellColls[element.sheet + "!" + element.cell];
      }
      Utils.UpdateCellValue(element.sheet, element.cell, result);
    });

    return Utils.getDocumentJson();
  }
}
