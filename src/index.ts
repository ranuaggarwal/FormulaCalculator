import { DefaultFormulaEngine } from "./FormulaEngine/DefaultFormulaEngine";
import * as sampleSolutionJson from "./Samples/Solution";

export class FormulaCalculator {
  public Evaluate(JSON: any): any {
    //For debugging
    if (JSON === undefined || JSON.length === 0) {
      JSON = sampleSolutionJson.solutionJson;
    }

    //Initialize formula engine
    const formulaEngine = new DefaultFormulaEngine();

    try {
      JSON = formulaEngine.EvaluateFormula(JSON);
      return JSON;
    } catch (e) {
      console.log("Error Returning calculating formula. Error: " + e);
      return JSON;
    }
  }
}

const formulaCalculator: FormulaCalculator = new FormulaCalculator();
formulaCalculator.Evaluate("");
