/* tslint:disable:prefer-for-of */
/* tslint:disable*/

import * as formulaParser from "libs-leonardo-formula-parser";
import { Utils } from "../Common/Utils";

export class FormulaUtils {
  static EvaluateFormula(
    formula: string,
    sheet: string,
    InputCellColls: any
  ): any {
    var parser = new formulaParser.Parser();
    parser.setVariable("true", 1);
    parser.setVariable("false", 0);

    parser.on("callCellValue", function(cellCoord, done) {
      var sheetName;
      if (cellCoord.sheet == undefined) {
        sheetName = sheet;
      } else {
        sheetName = cellCoord.sheet;
      }
      var row = cellCoord.row.index;
      var col = cellCoord.column.index;

      var cellAlias =
        sheetName + "!" + cellCoord.column.label + cellCoord.row.label;

      if (InputCellColls[cellAlias] != null) {
        done(InputCellColls[cellAlias]);
      } else {
        var dependentCellFormula = Utils.GetFormulaIfFormulaCell(
          cellCoord.column.label + cellCoord.row.label,
          sheetName
        );
        if (dependentCellFormula != null) {
          var result = parser.parse(dependentCellFormula).result;
          if (!InputCellColls[cellAlias]) {
            InputCellColls[cellAlias] = result;
          }

          done(result);
        } else {
          var matchedCellValue = Utils.GetValueOfNodeFromDocumentn(
            row,
            col,
            sheetName
          );
          done(matchedCellValue);
        }
      }
    });

    parser.on("callRangeValue", function(startCellCoord, endCellCoord, done) {
      var fragment = [];
      var sheetName = startCellCoord.sheet;
      for (
        var row = startCellCoord.row.index;
        row <= endCellCoord.row.index;
        row++
      ) {
        var colFragment = [];

        for (
          var col = startCellCoord.column.index;
          col <= endCellCoord.column.index;
          col++
        ) {
          var value = 0;
          var cellAlias = sheetName + "!" + Utils.GetCol(col) + (row + 1);
          if (InputCellColls[cellAlias] != null) {
            value = InputCellColls[cellAlias];
          } else {
            var dependentCellFormula = Utils.GetFormulaIfFormulaCell(
              Utils.GetCol(col) + (row + 1),
              sheetName
            );
            if (dependentCellFormula != null) {
              value = parser.parse(dependentCellFormula).result;
              if (!InputCellColls[cellAlias]) {
                InputCellColls[cellAlias] = value;
              }
            } else {
              value = Utils.GetValueOfNodeFromDocumentn(row, col, sheetName);
            }
          }

          colFragment.push(value);
        }
        fragment.push(colFragment);
      }

      if (fragment) {
        done(fragment);
      }
    });

    try {
      return parser.parse(formula).result;
    } catch (error) {
      return 0;
    }
  }
}
