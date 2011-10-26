using System;
using System.Collections;
using System.Data;
using DynamicEquationEvaluator;
using MetricManagerClasses;
using MinorProcesses;
using DatabaseAccess;
using System.Diagnostics;

namespace EarningGeneratorClasses
{

...

			// get equation ids and instantiate Evaluators for each of them
			foreach( DataRow drEquation in dtEquations.Rows )
			{
				Decimal	decEquationId	= Convert.ToDecimal( drEquation[ "EQUATION_ID" ] );

				// initialize the object that will compile the equation in memory
				EquationProcess	procValidateEquation	= 
					new EquationProcess( dbAccess, 0M, decEquationId );

				// compile
				Evaluator	eEvaluator	= procValidateEquation.GetCompiledEquation();

				procValidateEquation.FinalizeProcess();

				// the evaluator will call the metric manager himself
				// in order to fetch the needed metrics
				eEvaluator.SetMetricManager( mmMetricManager );

				// add the equation DLL indexed by its id
				evalList.Add( decEquationId, eEvaluator );
			}
...
}
