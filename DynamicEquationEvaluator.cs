using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using MetricManagerClasses;
using System.IO;
using System.Text.RegularExpressions;

namespace DynamicEquationEvaluator
{
	/// <summary>
	/// Class that validates equations, creates and instantiate equation DLLs
	/// </summary>
	[Serializable]
	public class Evaluator
	{
		// tells the DynamicEquationEvaluator to compile equation with tracing steps
		public bool	bDebug	= false;

		#region Atributes

		// constant used as the function name by the static Evaluate methods
		const string strStaticMethodName = "__foo";
		// name of the last loaded DLL
		string strLoadedDLLName	= "";
		// object with the instantiated dynamic class
		object _objInstantiatedEvaluator = null;
		// app domain is used so we can unload the DLL when it must be replaced
		AppDomain adEvaluatorDomain;//	= AppDomain.CreateDomain( "EvaluatorDomain" );
		// path where the Metric Manager dll resides
		string strMetricManagerPath	= "MetricManager.dll";
		// list of operators
		public static char[]	caOPERATOR_LIST	= "+-*/^<>%?:=".ToCharArray();

		#endregion

		#region Public methods that evaluate the equations already set up in the constructor

		// all these typed Evaluate methods will call the main Evaluate method
		// if needed, additional typed Evaluate methods can be created
		public int EvaluateInt( string strEquationName )
		{
			return ( int ) Evaluate( strEquationName );
		}

		public string EvaluateString( string strEquationName )
		{
			return ( string ) Evaluate( strEquationName );
		}

		public bool EvaluateBool( string strEquationName )
		{
			return ( bool ) Evaluate( strEquationName );
		}

		public Decimal EvaluateDecimal( string strEquationName )
		{
			return ( Decimal ) Evaluate( strEquationName );
		}

		// this is the main Evaluate method	that returns the result cast as object
		public object Evaluate( string strEquationName )
		{
			strEquationName	= ConvertEquationName( strEquationName );

			// get the method information by its name
			MethodInfo	miMethod	= _objInstantiatedEvaluator.GetType().GetMethod( 
				"get" + strEquationName );

			if( miMethod == null )
			{
				// if method was not found, the unload the DLL before reporting error
				this.FreeDLL();

				throw new Exception( "Could not find method get" + 
					strEquationName + " in " + strLoadedDLLName );
			}

			//				// for debugging purposes
			//				StreamWriter	swSourceFile	= 
			//					new StreamWriter( File.OpenWrite( @"\emcompdev\log\Evaluate_DEBUG.TXT" ) );
			//				swSourceFile.Write( strEquationName + " " + strLoadedDLLName + "  " );
			//				swSourceFile.Close();


			// execute the method. the result is cast as object
			// and later it will be recast to its specified type
			return	miMethod.Invoke( _objInstantiatedEvaluator, null );
		}

		// this is the Evaluate methodthat returns the Qty Source result cast as Decimal
		public Decimal EvaluateQtySource( string strEquationName )
		{
			strEquationName	= ConvertEquationName( strEquationName );

			// get the method information by its name
			MethodInfo	miMethod	= _objInstantiatedEvaluator.GetType().GetMethod( 
				"QuantityOf" + strEquationName );

			if( miMethod == null )
			{
				// if method was not found, the unload the DLL before reporting error
				this.FreeDLL();

				return 0m;
//				throw new Exception( "Could not find method QuantityOf" + 
//					strEquationName + " in " + strLoadedDLLName );
			}

			//				// for debugging purposes
			//				StreamWriter	swSourceFile	= 
			//					new StreamWriter( File.OpenWrite( @"\emcompdev\log\Evaluate_DEBUG.TXT" ) );
			//				swSourceFile.Write( strEquationName + " " + strLoadedDLLName + "  " );
			//				swSourceFile.Close();


			// execute the method. the result is cast as object
			// and later it will be recast to its specified type
			return (Decimal) miMethod.Invoke( _objInstantiatedEvaluator, null );
		}

		#endregion

		#region Static methods for compiling and running C# code passed in a string

		// static methods that evaluate the C# code passed
		// if needed, additional typed Evaluate methods can be created
		static public int EvaluateToInteger( string strCSharpCode )
		{
			Evaluator eval	= new Evaluator( typeof( int ), 
				strCSharpCode, 
				strStaticMethodName );
			return ( int ) eval.Evaluate( strStaticMethodName );
		}

		static public string EvaluateToString( string strCSharpCode )
		{
			Evaluator eval	= new Evaluator( typeof( string ),
				strCSharpCode, 
				strStaticMethodName );
			return ( string ) eval.Evaluate( strStaticMethodName );
		}

		static public bool EvaluateToBool( string strCSharpCode )
		{
			Evaluator eval	= new Evaluator( typeof( bool ), 
				strCSharpCode, 
				strStaticMethodName );
			return ( bool ) eval.Evaluate( strStaticMethodName );
		}

		static public Decimal EvaluateToDecimal( string strCSharpCode )
		{
			Evaluator eval	= new Evaluator( typeof( Decimal ), 
				strCSharpCode, 
				strStaticMethodName );
			return ( Decimal ) eval.Evaluate( strStaticMethodName );
		}

		static public object EvaluateToObject( string strCSharpCode )
		{
			// this will compile the code passed as argument inside a function 
			// named __foo and then execute the function and return the result
			Evaluator eval	= new Evaluator( typeof( object ), 
				strCSharpCode, 
				strStaticMethodName );
			return eval.Evaluate( strStaticMethodName );
		}

		#endregion

		#region Public Load method to load dynamic class assembly from a DLL

		public bool EquationDLLExists( string strDLLFileNameParam )
		{
			// check if DLL exists
			FileInfo fiDLL	= new FileInfo( strDLLFileNameParam );
			
			return fiDLL.Exists;
		}

		public bool LoadEquationsFromDLL( string strDLLFileNameParam )
		{
			// if the DLL is already loaded, return immediately
			if( strLoadedDLLName.Equals( strDLLFileNameParam ) )
				return true;

			// if there is a DLL loaded, the application domain where it resides
			// must be unloaded (releasing any DLL previously loaded) before 
			// a new DLL is loaded
			if( !strLoadedDLLName.Equals( "" ) ) 
				FreeDLL();

			// check if DLL exists
			if( !EquationDLLExists( strDLLFileNameParam ) )
				return false;
			//throw new Exception( "Could not find DLL " + fiDLL.FullName );

			// store the new DLL name
			strLoadedDLLName	= strDLLFileNameParam;

			// reinstantiate app domain to load a new DLL
			adEvaluatorDomain	= AppDomain.CreateDomain( "EvaluatorDomain" );

			try
			{
				// create instance of evaluator within the application domain
				_objInstantiatedEvaluator		= adEvaluatorDomain.CreateInstanceFromAndUnwrap( 
					strDLLFileNameParam, "ExpressionEvaluator._Evaluator" );
			}
			catch( BadImageFormatException eBadFormat )
			{
				throw eBadFormat;
			}

			return true;

			//			//Load an assembly with the specified dll 
			//			Assembly assemblyEvaluatorClass	= Assembly.LoadFrom( strDLLFileNameParam );
			//			
			//			_objInstantiatedEvaluator		= assemblyEvaluatorClass.CreateInstance( 
			//											"ExpressionEvaluator._Evaluator" );
			//			return true;

		}


		#endregion

		#region Public methods to change variable values in the dynamic class

		public void ChangeVariable( string strVariableName, Decimal decValue )
		{
			// get the variable (attribute) by name
			FieldInfo fiTest	= 
				_objInstantiatedEvaluator.GetType().GetField( strVariableName );

			if( fiTest == null )
				throw new Exception( "Could not find attribute " + 
					strVariableName + " in " + strLoadedDLLName );

			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decValue );

		}

		public bool VariableExists( string strVariableName )
		{
			FieldInfo fiVariable	= 
				_objInstantiatedEvaluator.GetType().GetField( strVariableName );

			if( fiVariable == null )
				return false;

			return true;
		}


		public void SetMetricManager( MetricManager mmMetricManager ) 
		{
			// get the variable (attribute) by name
			FieldInfo fiTest	= 
				_objInstantiatedEvaluator.GetType().GetField( "mmMetricManager" );

			if( fiTest == null )
				throw new Exception( "Could not find attribute mmMetricManager" + 
					" in " + strLoadedDLLName );

			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) mmMetricManager );
		}


		public void SetParameters( DateTime dtBegin, DateTime dtEnd, 
			Decimal decProvLocId, Decimal decLocationId, 
			Decimal decProcessId, Decimal decPayHeaderId, 
			String	strScheduleType ) 
		{
			// get the variable (attribute) by name
			FieldInfo fiTest	= 
				_objInstantiatedEvaluator.GetType().GetField( "dtBegin" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) dtBegin );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "dtEnd" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) dtEnd );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decProvLocId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decProvLocId );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decLocationId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decLocationId );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decProcessId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decProcessId );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decPayHeaderId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decPayHeaderId );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "strScheduleType" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) strScheduleType );
		}


		public void SetParameters( DateTime dtBegin, DateTime dtEnd, 
			Decimal decProvLocId, Decimal decProcessId ) 
		{
			// get the variable (attribute) by name
			FieldInfo fiTest	= 
				_objInstantiatedEvaluator.GetType().GetField( "dtBegin" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) dtBegin );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "dtEnd" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) dtEnd );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decProvLocId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decProvLocId );

			// get the variable (attribute) by name
			fiTest	= _objInstantiatedEvaluator.GetType().GetField( "decProcessId" );
			// change the value of the variable
			fiTest.SetValue( _objInstantiatedEvaluator, (object) decProcessId );
		}


		#endregion

		#region External API function declarations to release DLLs

		[DllImport( @"kernel32" )]
		private extern static void FreeLibrary( IntPtr hLibModule ); 
		// cannot get handle to the DLL yet, so this method is not yet usable

		[DllImport( @"kernel32" )]
		private extern static IntPtr GetModuleHandle( string strModuleName );

		#endregion

		#region Public method that compiles the _Evaluator class 

		// method that dynamically compiles and instantiates the _Evaluator class
		public void ConstructEvaluator( EquationDefinition[] edEquations, VariableDefinition[] vdVariables, string strDLLFileNameParam )
		{
			// instantiate a C# Code Compiler and a Compiler Parameter to use with it
			//ICodeCompiler CSharpCodeCompiler = (new CSharpCodeProvider().CreateCompiler());
			CodeDomProvider CSharpCodeCompiler = CodeDomProvider.CreateProvider("CSharp");
			
			CompilerParameters	cpCompilerParameters	= new CompilerParameters();

			// the Parameter object will tell the compiler what DLLs to use
			cpCompilerParameters.ReferencedAssemblies.Add( "system.dll" );

			if( bDebug )
				AddTracingAssemblies( cpCompilerParameters );

			// add Metric Manager reference
			String	strPath	= System.Configuration.ConfigurationManager.AppSettings[ "AppPath" ].ToString();
			cpCompilerParameters.ReferencedAssemblies.Add( strPath + strMetricManagerPath );

			if( strDLLFileNameParam.Equals( "" ) )
				// this tells to compile the code (in memory only) and to not generate .EXE
				cpCompilerParameters.GenerateExecutable	= false;
			else 
			{
				// unload the app domain with the DLL
				FreeDLL();

				// release the NEW DLL if it is loaded (who knows? it could be)
				// because we're going to overwrite it
				FreeLibrary( GetModuleHandle( strDLLFileNameParam ) );

				// deletes the file because we're going to create 
				// a new one with this name
				File.Delete( strDLLFileNameParam );

				// this tells to generate executable, that is, a DLL
				cpCompilerParameters.GenerateExecutable	= true;
				cpCompilerParameters.OutputAssembly		= strDLLFileNameParam;

				strLoadedDLLName	= strDLLFileNameParam;
			}

			cpCompilerParameters.GenerateInMemory	= true;

			// create a string builder that will 
			// hold the source code in C#
			StringBuilder strbCSharpCode	= new StringBuilder();

			if( bDebug )
			{
				// these are needed to save Xml from inside the equation
				strbCSharpCode.Append( @"
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Xml;
using System.IO;
" );
			}

			// write the C# header in the string builder.
			// inside of the C# code there will be a class _Evaluator
			// the _Evaluator class will have functions that will return the result
			// of the expressions
			// the use of adEvaluatorDomain requires compiled assembly to be serializable
			// add attributes to hold MetricManager instance and some parameters
			// and add methods GreaterOf and LesserOf
			strbCSharpCode.Append( @"
using System;
using MetricManagerClasses;

namespace ExpressionEvaluator {

	[Serializable]
	public class _Evaluator {
		public MetricManager mmMetricManager;
		public DateTime dtBegin;
		public DateTime dtEnd;
		public Decimal decLocationId;
		public Decimal decProvLocId;
		public Decimal decProcessId;
		public Decimal decPayHeaderId;
		public String strScheduleType;

" );

			// add methods for > and <
			if( CheckExpression( edEquations, "GreaterOf(" ) )
				strbCSharpCode.Append( @"
		public Decimal GreaterOf( Decimal decValue1, Decimal decValue2 )
		{
			return decValue1 > decValue2 ? decValue1 : decValue2;
		}

" );
			if( CheckExpression( edEquations, "LesserOf(" ) )
				strbCSharpCode.Append( @"
		public Decimal LesserOf( Decimal decValue1, Decimal decValue2 )
		{
			return decValue1 < decValue2 ? decValue1 : decValue2;
		}

" );


			// append constant definitions that will store the vdVariable values
			foreach( VariableDefinition vdVariable in vdVariables ) 
			{
				if( Convert.IsDBNull( vdVariable ) || vdVariable == null )
					continue;

				if( !vdVariable.Cacheable ) 
				{
					// create public vdVariable in the format:  public Decimal NAME = VALUE;
					strbCSharpCode.AppendFormat( "		public Decimal {0} = {1};\r\n", 
						vdVariable.Name, vdVariable.Value );
				}
				else
				{
					// create a getter alias for every variable in order to call the 
					// MetricManager
					// special trick "\"" below is to insert double quotes 
					// inside the StringBuilder
					strbCSharpCode.AppendFormat( "\r\n		public Decimal {0}\r\n", vdVariable.Name );
					strbCSharpCode.Append( "		{ get { return mmMetricManager.GetMetricValue( \"" );
					strbCSharpCode.Append( vdVariable.Name );
					strbCSharpCode.Append("\", \r\n" );
					strbCSharpCode.Append( @"
					dtBegin, dtEnd, decProvLocId, 
					decLocationId, decProcessId, 
					decPayHeaderId, strScheduleType ); } }
" );
				}
			}

			// skip a line and add main method to satisfy the compiler 
			// in case of DLL generation required
			strbCSharpCode.Append( "\r\n" );
			if( !strDLLFileNameParam.Equals( "" ) )
			{
				strbCSharpCode.Append( @"
		public static void Main() {}

" );
			}

			// append C# function definitions that will return the result of the expressions
			foreach( EquationDefinition edEquation in edEquations )
			{
				String	strEquationName	= ConvertEquationName( edEquation.Name );

				// create public function in the format:  public TYPE getNAME()
				strbCSharpCode.AppendFormat( "\r\n		public {0} get{1}()\r\n	", edEquation.ReturnType.Name, strEquationName );
				strbCSharpCode.Append( "{ \r\n" );

				// define the function in the format:  { return ( EXPRESSION ); }
				// or { return ( CONDITION ? EXPRESSION : 0 ); } if there is a condition
				if( edEquation.ConditionExpression.Equals( "" ) )
				{
					if( bDebug )
						AddTracingCode( strbCSharpCode );

					// if there is division, enclose expression in a try-catch
					if( edEquation.Expression.IndexOf( "/" ) > 0 )
					{
						strbCSharpCode.Append( "		try { \r\n      return ( " );
						strbCSharpCode.Append( edEquation.Expression );
						strbCSharpCode.Append( " ); \r\n		} \r\n		catch( Exception excpt ) {\r\n" );
						strbCSharpCode.Append( "		if( excpt.Message.IndexOf( \"Attempted to divide by zero\" ) >= 0 ) return 0; \r\n" );
						strbCSharpCode.Append( "			throw( excpt ); \r\n" );
						strbCSharpCode.Append( "		} \r\n" );
					}
					else
						strbCSharpCode.AppendFormat( "		return ( {0} ); \r\n", edEquation.Expression );

				}
				else
				{
					// if there is division, enclose expression in a try-catch
					if( edEquation.Expression.IndexOf( "/" ) > 0 )
					{
						strbCSharpCode.Append( "		try { \r\n      " );
						strbCSharpCode.AppendFormat( "return ( {0} ? {1} : 0m ); \r\n", 
							edEquation.ConditionExpression, edEquation.Expression );
						strbCSharpCode.Append( "		} \r\n		catch( Exception excpt ) {\r\n" );
						strbCSharpCode.Append( "		if( excpt.Message.IndexOf( \"Attempted to divide by zero\" ) >= 0 ) return 0; \r\n" );
						strbCSharpCode.Append( "			throw( excpt ); \r\n" );
						strbCSharpCode.Append( "		} \r\n" );
					}
					else
						strbCSharpCode.AppendFormat( "		return ( {0} ? {1} : 0m ); \r\n", 
							edEquation.ConditionExpression, edEquation.Expression );

				}

				strbCSharpCode.Append( "}\r\n\r\n" );

				// create a getter alias for every equation name (in order to not use 
				// parenthesis when referencing the equation function)
				strbCSharpCode.AppendFormat( "		public {0} {1} \r\n", 
					edEquation.ReturnType.Name, strEquationName );
				strbCSharpCode.Append( "		{ get { return " );
				strbCSharpCode.AppendFormat( "get{0}", strEquationName );
				strbCSharpCode.Append( "(); } }\r\n\r\n" );

				// if there is a Qty Source specified, create a 
				// public function in the format:  public TYPE QuantityOfNAME()
				if( ! edEquation.QuantitySource.Equals( "" ) )
				{
					strbCSharpCode.AppendFormat( "\r\n		public {0} QuantityOf{1}",
								edEquation.ReturnType.Name, strEquationName );
					strbCSharpCode.Append( "()\r\n		{ \r\n" );
					
					// if there is a condition, arrange it so it only calculates quantity if condition is evaluated as true
					if( edEquation.ConditionExpression.Equals( "" ) )
						// define the function in the format:  { return ( QTY EXPRESSION ); }
						strbCSharpCode.AppendFormat( "		return ( {0} ); \r\n",
									edEquation.QuantitySource );
					else
						// define the function in the format:  { return ( condition ? QTY EXPRESSION : 0m ); }
						strbCSharpCode.AppendFormat( "return ( {0} ? {1} : 0m ); \r\n",
							edEquation.ConditionExpression, edEquation.QuantitySource );

					strbCSharpCode.Append( "		}\r\n\r\n" );
				}
			}

			// close the C# code
			strbCSharpCode.Append( "} }" );

			// save source code if debugging (uncomment code in the method below)
			if( bDebug )
				SaveSourceCode( strLoadedDLLName, strbCSharpCode );

			// send the C# code in the string builder to the compiler and get the results
			CompilerResults crCompilerResults	= 
				CSharpCodeCompiler.CompileAssemblyFromSource( 
				cpCompilerParameters, strbCSharpCode.ToString() );
			if( crCompilerResults.Errors.HasErrors )
			{
				if( strLoadedDLLName.Equals( "" ) )
				{
					String	strEQUATION_DLL_PREFIX	= 
						System.Configuration.ConfigurationManager.AppSettings[ 
						"DLLPathAndName" ];
					strLoadedDLLName	= strEQUATION_DLL_PREFIX + ConvertEquationName( edEquations[ 0 ].Name );
				}

				// save source code for debugging purposes
				StreamWriter	swSourceFile	= new StreamWriter( File.OpenWrite( strLoadedDLLName + "_with_error.cs" ) );
				swSourceFile.Write( strbCSharpCode.ToString() );
				swSourceFile.Close();

				// create a string builder listing all errors and throw an exception
				StringBuilder	strbErrors	= new StringBuilder();
				strbErrors.Append( "Error Compiling Expression: " );
				foreach( CompilerError err in crCompilerResults.Errors )
				{
					strbErrors.AppendFormat( "{0}\r\n(Error {1}, Line {2}, Column {3})", 
						err.ErrorText, err.ErrorNumber,
						err.Line.ToString(), err.Column.ToString() );
				}
				throw new Exception( strbErrors.ToString() );
			}

			// reinstantiate app domain and load the new DLL
			adEvaluatorDomain	= AppDomain.CreateDomain( "EvaluatorDomain" );

			try
			{
				// get the dynamically compiled assembly and create an instance 
				// of the _Evaluator class
				Assembly	assemblyEvaluatorClass		= crCompilerResults.CompiledAssembly;

				_objInstantiatedEvaluator	= assemblyEvaluatorClass.CreateInstance( 
					"ExpressionEvaluator._Evaluator" );

				//				// create instance of evaluator (from the DLL) within the application domain
				//				_objInstantiatedEvaluator		= adEvaluatorDomain.CreateInstanceFromAndUnwrap( 
				//					strDLLFileNameParam, "ExpressionEvaluator._Evaluator" );
			}
			catch( BadImageFormatException eBadFormat )
			{
				throw eBadFormat;
			}

		}

		private bool CheckExpression( EquationDefinition[] edEquations, String strExpression )
		{
			foreach( EquationDefinition objEquation in edEquations )
				if( objEquation.ConditionExpression.IndexOf( strExpression ) >= 0 
					|| objEquation.Expression.IndexOf( strExpression ) >= 0 )
					return true;

			return false;
		}

		private void AddTracingAssemblies( CompilerParameters cpParam )
		{
			// these are needed to save Xml from inside the equation
			cpParam.ReferencedAssemblies.Add( "system.data.dll" );
			cpParam.ReferencedAssemblies.Add( "system.xml.dll" );
		}
		private void AddTracingCode( StringBuilder strbCSharpCodeParam )
		{
			strbCSharpCodeParam.Append( @"
// for debugging purposes
StreamWriter	swSourceFile	= 
		new StreamWriter( File.OpenWrite( @" + "\"\\emcompdev\\log\\DLL_DEBUG.TXT\" ) );" + @"
swSourceFile.Write( " + "\"GetMetricValue is being called: \" + mmMetricManager.ToString() + \".\" );" + @"
//swSourceFile.Write( " + "\"TotalHoursWorked: \" + TotalHoursWorked.ToString() + \" . \" );" + @"
//swSourceFile.Write( " + "\"PatientsPerHourFactor: \" + PatientsPerHourFactor.ToString() + \" . \" );" + @"
//swSourceFile.Write( " + "\"Result: \" + ( TotalHoursWorked * PatientsPerHourFactor ).ToString() + \" . \" );" + @"
swSourceFile.Close();
" );
		}
		private void SaveSourceCode( String strLoadedDLLNameParam, StringBuilder strbCSharpCodeParam )
		{
			// save source code for debugging purposes
			StreamWriter	swSourceFile1	= 
				new StreamWriter( File.OpenWrite( 
				strLoadedDLLNameParam + "_source.cs" ) );
			swSourceFile1.Write( strbCSharpCodeParam.ToString() );
			swSourceFile1.Close();
			return;
		}


		#endregion

		#region Methods to instantiate the evaluator, free the DLL, validate expressions

		/// <summary>
		/// unload app domain with the DLL
		/// </summary>
		public void FreeDLL()
		{
			// unload the app domain with the DLL previously loaded
			_objInstantiatedEvaluator	= null;

			if( adEvaluatorDomain != null )
				AppDomain.Unload( adEvaluatorDomain );

			adEvaluatorDomain	= null;

			// release the new DLL if it is loaded (could it be?)
			if( !this.strLoadedDLLName.Equals( "" ) )
				FreeLibrary( GetModuleHandle( this.strLoadedDLLName ) );


			//			// release the DLL since it is not being used anymore
			//			IntPtr	ipModule	= GetModuleHandle( strLoadedDLLName );
			//			if( ipModule.ToInt32() > 0 )
			//				FreeLibrary( ipModule );

		}


		/// <summary>
		/// Returns valid identifier string (with no invalid characters and not starting with a number)
		/// </summary>
		public static String ConvertEquationName( String strEqNameParam )
		{
			char[]	caInvalidChrs	= 
				" ~`!@#$%^&*()_-+=\"\\|[]{};:',.<>/?".ToCharArray();

			// eliminate spaces and invalid characters
			foreach( char cInvalid in caInvalidChrs )
				strEqNameParam	= strEqNameParam.Replace( cInvalid.ToString(), "" );

			// make sure it does not start with a number
			int iPos	= strEqNameParam.IndexOfAny( "0123456789".ToCharArray() );
			if( iPos.Equals( 0 ) )
				strEqNameParam	= "Number" + strEqNameParam;

			return strEqNameParam;
		}

		public static String FixQtySource( String strQtySource )
		{
			if( strQtySource.Equals( "" ) )
				return "";

			// ignore numbers, parenthesis and minus sign
			if( strQtySource.IndexOfAny( "0123456789(-".ToCharArray() ) >= 0 )
				return strQtySource;

			// ignore functions
			if( strQtySource.IndexOf( "GreaterOf" ) >= 0 )
				return strQtySource;
			if( strQtySource.IndexOf( "LesserOf" ) >= 0 )
				return strQtySource;

			// add square brackets to metric name if it does not already have []
			strQtySource	= strQtySource.Trim();
			if( ! strQtySource.Substring( 0, 1 ).Equals( "[" ) )
				strQtySource	= String.Format( "[{0}]", strQtySource );

			return strQtySource;
		}


		/// <summary>
		/// Returns the equation text with no duplicate, leading or trailing spaces
		/// </summary>
		public static String RemoveExtraSpaces( String strEquationTxt )
		{
			// eliminate carriage return and line breaks
			strEquationTxt	= strEquationTxt.Replace( "\r", " " );
			strEquationTxt	= strEquationTxt.Replace( "\n", " " );

			// remove all consecutive spaces
			while( strEquationTxt.IndexOf( "  " ) > 0 )
				strEquationTxt	= strEquationTxt.Replace( "  ", " " );

			// remove leading or trailing spaces
			strEquationTxt	= strEquationTxt.Trim();

			return strEquationTxt;
		}


		/// <summary>
		/// Returns the equation text with literals appended "M" to indicate a Decimal literal
		/// </summary>
		public static String ConvertLiteralsToDecimal( String strEquationTxt )
		{
			char[]	caNumbers	= "0123456789".ToCharArray();
			int	iPos	= strEquationTxt.IndexOfAny( caNumbers );
			while( iPos >= 0 )
			{
				// if found number at the end of equation, append "M" to it
				// and finish
				if( iPos.Equals( strEquationTxt.Length - 1 ) )
				{
					strEquationTxt	= strEquationTxt + "M";
					break;
				}

				// if found space or operators after a number, append "M" to it
				if( strEquationTxt.Substring( iPos + 1, 1 ).Equals( " " ) )
				{
					// try to detect if this number is a variable name (has closing 
					// brackets) or a literal (is outside of any brackets)
					int	iPosClosing	= strEquationTxt.IndexOf( "]", iPos + 1 );
					if( iPosClosing < 0 )
						iPosClosing	= strEquationTxt.Length;
					int	iPosOpening	= strEquationTxt.IndexOf( "[", iPos + 1 );
					if( iPosOpening < 0 )
						iPosOpening	= strEquationTxt.Length;
					// if closing bracket is preceeded by an opening bracket
					// or there is no closing bracket, then this is a literal
					if( iPosClosing >= iPosOpening ) 
						strEquationTxt	= strEquationTxt.Insert( iPos + 1, "M" );
				}

				foreach( char caOperator in caOPERATOR_LIST )
				{
					if( strEquationTxt.Substring( iPos + 1, 1 ).Equals( caOperator ) )
					{
						strEquationTxt	= strEquationTxt.Insert( iPos + 1, "M" );
						break;
					}
				}

				// search for the next number
				iPos	= strEquationTxt.IndexOfAny( caNumbers, iPos + 1 );
			}

			return strEquationTxt;
		}


		/// <summary>
		/// Check if there is an unequal number of parenthesis/brackets/quotes/etc
		/// </summary>
		public static void CheckBalancing( String strEnclosingChrs, 
			String strParam, String strConditionOrEquation )
		{
			// check parenthesis/brackets balancing
			char[]	caEnclosing	= strEnclosingChrs.ToCharArray();
			int	iPosFirst	= strParam.IndexOfAny( caEnclosing );
			int iPos	= iPosFirst;
			int iLevel	= 0;
			while( iPos >= 0 )
			{
				// increase level when there is an opening parenthesis
				if( strParam.Substring( iPos, 1 ).Equals( 
					caEnclosing[ 0 ].ToString() ) )
					iLevel ++; 
				// decrease level when there is a closing parenthesis
				if( strParam.Substring( iPos, 1 ).Equals( 
					caEnclosing[ 1 ].ToString() ) )
					iLevel --; 

				// if level is negative, there is a closing parenthesis not needed
				if( iLevel < 0 )
					throw new Exception( strConditionOrEquation + " Validation Error:  Incorrect syntax near '" +
						caEnclosing[ 1 ].ToString() + 
						"'.\r\n[POSITION:" + iPos.ToString() + "]" );

				// look for next parenthesis
				iPos	= strParam.IndexOfAny( caEnclosing, iPos + 1, 
					strParam.Length - iPos - 1 );
			}

			// if level ended up positive, there is an opening parenthesis not needed
			if( iLevel > 0 )
				throw new Exception( strConditionOrEquation + " Validation Error:  Incorrect syntax near '" +
					caEnclosing[ 0 ].ToString() + 
					"'.\r\n[POSITION:" + iPosFirst.ToString() + "]" );
		}


		/// <summary>
		/// Returns string with a list of metrics in quotes and separated by commas
		/// </summary>
		public static StringBuilder GetMetricList( String strEquationTxt )
		{
			StringBuilder	strbMetrics	= new StringBuilder( "'" );
			String	strRegex	= @"(?<PreOperator>[^\][]*)\[(?<Metric>[^\]]*)](?<Operator>[^\][]*)";
			Regex	objRegex	= new Regex( strRegex );
			MatchCollection	objMatches	= objRegex.Matches( strEquationTxt );

			// collect all metrics in between square brackets
			foreach( Match objMatch in objMatches )
			{
				// collect metric and append to the list
				String	strMetric	= objMatch.Groups[ "Metric" ].Value;
				if( !strMetric.Equals( "" ) )
					strbMetrics.AppendFormat( "{0}', '", strMetric );

				// check the preceding operator if any
				String	strOperator	= objMatch.Groups[ "PreOperator" ].Value;
				int	iPosOper;
				if( !strOperator.Equals( "" ) )
				{
					iPosOper	= strOperator.IndexOfAny( caOPERATOR_LIST );
					// if operator not found, check for opening parenthesis
					if( iPosOper < 0 )
					{
						// if no operator, neither parenthesis, return error
						if( strOperator.IndexOf( "(" ) < 0 )
							throw new Exception( 
								String.Format( "Validation Error:  Missing operator before metric '{0}'.\r\n[POSITION:{1}]", 
								strMetric, objMatch.Groups[ "Operator" ].Index.ToString() ) );
					}
				}

				// skip null string found at the beginning/end of the equation
				strOperator	= objMatch.Groups[ "Operator" ].Value;
				if( strOperator.Equals( "" ) ) 
					if( objMatch.Groups[ "Operator" ].Index.Equals( 0 ) 
						|| objMatch.Groups[ "Operator" ].Index.Equals( strEquationTxt.Length ))
						continue;

				iPosOper	= strOperator.IndexOfAny( caOPERATOR_LIST );
				// if operator not found, return error
				if( iPosOper < 0 )
				{
					// if no operator, neither parenthesis, return error
					if( strOperator.IndexOf( ")" ) < 0 ) 
						throw new Exception( 
							String.Format( "Validation Error:  Missing operator after metric '{0}'.\r\n[POSITION:{1}]", 
							strMetric, objMatch.Groups[ "Operator" ].Index.ToString() ) );
				}
			}

			// remove last comma and space
			if( strbMetrics.Length < 3 )
				return new StringBuilder( "" );
			else
				strbMetrics.Remove( strbMetrics.Length - 3, 3 );

			return strbMetrics;

			//			// identify all metrics used in the equation by the opening bracket "["
			//			int	iPos		= strEquationTxt.IndexOf( "[", 0, strEquationTxt.Length );
			//
			//			// if no metric was found, returns empty
			//			if( iPos < 0 )
			//				return new StringBuilder( "" );
			//
			//			int	iPosNext	= strEquationTxt.IndexOf( "]", iPos, strEquationTxt.Length - iPos );
			//			int	iPosPrevious, iPosOperator;
			//			String	strMetricName;
			//			StringBuilder	strbMetricList	= new StringBuilder( "\"" );
			//			while( iPos >= 0 )
			//			{
			//				// add metric to the list separated by commas
			//				strMetricName	= strEquationTxt.Substring( iPos + 1, iPosNext - iPos - 1 );
			//				strbMetricList.AppendFormat( "{0}\", \"", strMetricName );
			//
			//				// find the next metric
			//				iPosPrevious	= iPosNext;
			//				iPos		= strEquationTxt.IndexOf( "[", iPosNext, strEquationTxt.Length - iPosNext - 1 );
			//
			//				// if there is a next metric, get its end position
			//				if( iPos > 0 )
			//				{
			//					iPosNext	= strEquationTxt.IndexOf( "]", iPos, strEquationTxt.Length - iPos );
			//
			//					// check if there is an operator between two metrics
			//					// find the first operator
			//					iPosOperator	= strEquationTxt.IndexOfAny( 
			//						caOPERATOR_LIST, 
			//						iPosPrevious + 1, iPos - iPosPrevious - 1 );
			//					// if operator not found, return error
			//					if( iPosOperator < 0 )
			//						throw new Exception( "Validation Error:  Missing operator after metric '" +
			//							strMetricName + "'.\r\n[POSITION:" + iPosPrevious.ToString() + "]" );
			//				}
			//			}
			//
			//			// remove last comma and space
			//			strbMetricList.Remove( strbMetricList.Length - 3, 3 );
			//
			//			return strbMetricList;
		}


		#endregion

		#region Several Constructors

		public Evaluator()
		{
			//			EquationDefinition[]	edEquations	= {};
			//			VariableDefinition[]	vdVariables	= {};
			//			ConstructEvaluator( edEquations, vdVariables, "" );
		}
		// constructor that receives an array of edEquations
		public Evaluator( EquationDefinition[] edEquations )
		{
			VariableDefinition[]	vdVariables	= {};
			ConstructEvaluator( edEquations, vdVariables, "" );
		}

		// constructor that receives return type, expression and name of the expression
		public Evaluator( Type returnType, string expression, string strEquationName )
		{
			VariableDefinition[]	vdVariables	= {};
			EquationDefinition[]	edEquations	= { new EquationDefinition( returnType, expression, strEquationName ) };
			ConstructEvaluator( edEquations, vdVariables, "" );
		}

		// constructor that receives just one EquationDefinition
		public Evaluator( EquationDefinition edEquation )
		{
			VariableDefinition[]	vdVariables	= {};
			EquationDefinition[] edEquations	= { edEquation };
			ConstructEvaluator( edEquations, vdVariables, "" );
		}

		// constructor that receives EquationDefinition and VariableDefinition
		public Evaluator( EquationDefinition[] edEquations, VariableDefinition[] vdVariables )
		{
			ConstructEvaluator( edEquations, vdVariables, "" );
		}

		// constructor that receives EquationDefinition, VariableDefinition and DLL file name
		public Evaluator( EquationDefinition[] edEquations, VariableDefinition[] vdVariables, string strDLLFileNameParam )
		{
			ConstructEvaluator( edEquations, vdVariables, strDLLFileNameParam );
		}

		// constructor that receives just one EquationDefinition and the DLL name
		public Evaluator( EquationDefinition edEquation, string strDLLFileNameParam )
		{
			VariableDefinition[]	vdVariables	= {};
			EquationDefinition[] edEquations	= { edEquation };
			ConstructEvaluator( edEquations, vdVariables, strDLLFileNameParam );
		}


		#endregion
	}

	#region Parameter Classes

	public class VariableDefinition
	{
		public string	Name;
		public Decimal	Value;
		public bool		Cacheable;
		public string	MetricType;

		// basic constructor that just stores name and value
		public VariableDefinition( string strVariableName, Decimal fValue, bool bCacheable, string strMetricType )
		{
			Name	= strVariableName;
			Value	= fValue;
			Cacheable	= bCacheable;
			MetricType	= strMetricType;
		}
	}

	public class EquationDefinition
	{
		public Type		ReturnType;
		public string	Name;
		public string	Expression;
		public string	ConditionExpression;
		public string	QuantitySource;

		// basic constructor that just stores the return type, expression and expression name
		public EquationDefinition( Type tReturnType, string strExpression, 
			string strEquationName, string strCondition, string strQuantitySource )
		{
			ReturnType	= tReturnType;
			Expression	= strExpression;
			Name		= strEquationName;
			ConditionExpression	= strCondition;
			QuantitySource		= strQuantitySource;
		}

		public EquationDefinition( Type tReturnType, string strExpression, 
			string strEquationName, string strCondition )
		{
			ReturnType	= tReturnType;
			Expression	= strExpression;
			Name		= strEquationName;
			ConditionExpression	= strCondition;
			QuantitySource		= "";
		}

		public EquationDefinition( Type tReturnType, string strExpression, 
			string strEquationName )
		{
			ReturnType	= tReturnType;
			Expression	= strExpression;
			Name		= strEquationName;
			ConditionExpression	= "";
			QuantitySource		= "";
		}
	}


	#endregion
}
