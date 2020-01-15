using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Web;

namespace OportunidadesDynamics365
{
    /// <summary>
    /// Esta clase representa la serie de pasos apilados que mapean la 
    /// ruta ordenada que siguió una excepción.
    /// </summary>
    public class CustomCallStack
    {
        #region CustomCallStack - Propiedades

        private CustomCallStackFrame[] oCallStackFrames = null;

        #endregion CustomCallStack - Propiedades

        #region CustomCallStack - Propiedades Procedimentales

        /// <summary>
        /// Retorna el largo de la pila (read-only)
        /// </summary>
        public int Length
        {
            get
            {
                return oCallStackFrames.Length;
            }
        }

        /// <summary>
        /// Retorna la pila, CustomCallStackFrame (read-only)
        /// </summary>
        public CustomCallStackFrame[] Frames
        {
            get
            {
                return oCallStackFrames;
            }
        }


        #endregion CustomCallStack - Propiedades Procedimentales

        #region CustomCallStack - Constructores
        /// <summary>
        /// Constructor por defecto.
        /// </summary>
        public CustomCallStack()
        {
        }
        #endregion CustomCallStack - Constructores

        #region CustomCallStack - Métodos Públicos

        /// <summary>
        /// Coloca (Pushes) un nuevo paso dentro de la pila.
        /// </summary>
        public void Push(string sourceClass, string sourceMethod)
        {
            if (oCallStackFrames == null)
            {
                //----------------------------------------------------------------
                //-- Inicializa el arreglo de pasos ordenados (Frames).
                //----------------------------------------------------------------
                oCallStackFrames = new CustomCallStackFrame[1];
            }
            else
            {
                //----------------------------------------------------------------
                //-- Crea un arreglo temporal de pasos (Frames), de un elemento más largo que el existente.
                //----------------------------------------------------------------
                CustomCallStackFrame[] oTempFrames = new CustomCallStackFrame[oCallStackFrames.Length + 1];
                //----------------------------------------------------------------
                //--- Copia los elementos existentes sobre el arreglo de pasos temporal.
                //----------------------------------------------------------------
                oCallStackFrames.CopyTo(oTempFrames, 0);
                //----------------------------------------------------------------
                //--- Redefine el arreglo de pasos original.
                //----------------------------------------------------------------
                oCallStackFrames = oTempFrames;
                //----------------------------------------------------------------
                //--- Elimina el arreglo temporal pues ya no lo necesita más.
                //----------------------------------------------------------------
                oTempFrames = null;
            }
            //----------------------------------------------------------------
            //--- Crea un nuevo paso (CustomCallStackFrame) y lo coloca dentro del 
            //--- arreglo ordenado de pasos.
            //----------------------------------------------------------------
            oCallStackFrames[oCallStackFrames.Length - 1] = new CustomCallStackFrame(sourceClass, sourceMethod);
        }


        #endregion CustomCallStack - Métodos Públicos
    }

    /// <summary>
    /// Esta clase representa un paso específico dentro de la pila antes 
    /// definida.
    /// </summary>
    public class CustomCallStackFrame
    {
        #region CustomCallStackFrame - Propiedades

        public string m_sourceClass;
        public string m_sourceMethod;

        #endregion CustomCallStackFrame - Propiedades

        #region CustomCallStackFrame - Constructores

        /// <summary>
        /// Constructor por defecto.
        /// </summary>
        public CustomCallStackFrame()
        {
        }

        /// <summary>
        /// Constructor que setea en la instanciación, la clase y el método origen del 
        /// paso a ser apilado.
        /// </summary>
        public CustomCallStackFrame(string sourceClass, string sourceMethod)
        {
            this.m_sourceClass = sourceClass;
            this.m_sourceMethod = sourceMethod;
        }

        #endregion CustomCallStackFrame - Constructores

    }

    /// <summary>
    /// Descripción breve de NominaException.
    /// </summary>
    public class MsgException : ApplicationException
    {
        #region MsgException - Atributos

        private CustomCallStack m_CallStack = new CustomCallStack();
        private DateTime m_fechaOcurrencia = System.DateTime.Now;
        private string m_formatoFecha = "yyyy-MM-dd";
        private string m_EspaciadorDefault = " ---> ";
        private string m_EspaciadorHTMLDefault = "";
        private string m_EspaciadorHTMLStack = "";
        private string m_EspaciadorStack = " <--- ";
        private string m_storedProcedure = "";
        private int m_sqlErrorNumber = 0;
        private string m_tipoExcepcion = "";

        #endregion MsgException - Atributos

        #region MsgException - Propiedades Procedimientales

        /// <summary>
        /// La pila de la Excepción.
        /// </summary>
        public CustomCallStack CallStack
        {
            get
            {
                return this.m_CallStack;
            }
            set
            {
                this.m_CallStack = value;
            }
        }

        /// <summary>
        /// El tipo de excepción (read-only).
        /// </summary>
        public string Tipo
        {
            get
            {
                return this.m_tipoExcepcion;
            }
        }

        /// <summary>
        /// El stored procedure en el cual se originó la excepción (read-only).
        /// </summary>
        public string StoredProcedure
        {
            get
            {
                return this.m_storedProcedure;
            }
        }

        /// <summary>
        /// El número de error identificado por SQL Server, cuando se trata de 
        /// una excepción del tipo SqlException (read-only).
        /// </summary>
        public int SqlErrorNumber
        {
            get
            {
                return this.m_sqlErrorNumber;
            }
        }

        #endregion MsgException - Propiedades Procedimientales

        #region MsgException - Constructores

        /// <summary>
        /// Este método inicializa MsgException a través del seteo de algunas
        /// de sus propiedades.  En particular debe ser utilizado cuando se tiene una
        /// excepción básica como punto de partida.
        /// Utilizar cuando se va atrapa o lanzar una excepción estándar.
        /// </summary>
        public MsgException(string message, string sourceClassName, string sourceMethodName, Exception oInnerException)
            : base(message.Length == 0 ? oInnerException.Message : message, oInnerException)
        {
            InitializeSettings();

            //----------------------------------------------------------------
            //--- Setea el tipo de excepción.
            //----------------------------------------------------------------
            this.m_tipoExcepcion = "Exception";

            //----------------------------------------------------------------
            //--- Coloca la clase y método origen de la excepción en la pila.
            //----------------------------------------------------------------
            this.m_CallStack.Push(sourceClassName, sourceMethodName);
        }


        /// <summary>
        /// Este método inicializa MsgException a través del seteo de algunas
        /// de sus propiedades.  En particular debe ser utilizado cuando se tiene una
        /// excepción del tipo SqlException como punto de partida.
        /// Utilizar cuando se va atrapa o lanzar una excepción del tipo SqlException.
        /// </summary>
        public MsgException(string message, string sourceClassName, string sourceMethodName, SqlException oInnerException)
            : base(message.Length == 0 ? oInnerException.Message : message, oInnerException)
        {
            InitializeSettings();

            //----------------------------------------------------------------
            //--- Setea el tipo de excepción.
            //----------------------------------------------------------------
            this.m_tipoExcepcion = "SqlException";

            //----------------------------------------------------------------
            //--- Setea el nombre del stored procedure en el cual ocurrió la excepción.
            //----------------------------------------------------------------
            this.m_storedProcedure = oInnerException.Procedure;

            //----------------------------------------------------------------
            //--- Setea el número de error SQL de SqlException.
            //----------------------------------------------------------------
            this.m_sqlErrorNumber = oInnerException.Number;

            //----------------------------------------------------------------
            //--- Coloca la clase y método origen de la excepción en la pila.
            //----------------------------------------------------------------
            this.m_CallStack.Push(sourceClassName, sourceMethodName);
        }

        /// <summary>
        /// Este método inicializa MsgException a través del seteo de algunas
        /// de sus propiedades.  En particular debe ser utilizado cuando NO se tiene
        /// una excepción como punto de partida.
        /// Utilizar cuando se va atrapa o lanzar una excepción propia.
        /// </summary>
        public MsgException(string message, string sourceClassName, string sourceMethodName)
            : base(message)
        {
            InitializeSettings();

            //----------------------------------------------------------------
            //--- Setea el tipo de Exception.
            //----------------------------------------------------------------
            this.m_tipoExcepcion = "CustomException";

            //----------------------------------------------------------------
            //--- Coloca la clase y método origen de la excepción en la pila.
            //----------------------------------------------------------------
            this.m_CallStack.Push(sourceClassName, sourceMethodName);
        }

        #endregion MsgException - Constructores

        #region MsgException - Métodos Públicos

        /// <summary>
        /// Este método permite, de manera opcional, registrar el mensaje de la excepción 
        /// en un archivo o en el EventLog de Windows.
        /// </summary>
        public string Log(bool fileLog, bool eventLog)
        {
            //----------------------------------------------------------------
            //--- Coloca la cabecera del mensaje. 
            //----------------------------------------------------------------
            StringBuilder oSB = new StringBuilder();
            oSB.Append(this.m_fechaOcurrencia + " - " + "Ocurrió una Excepción del tipo: " + this.m_tipoExcepcion);

            //----------------------------------------------------------------
            //--- Las excepciones del tipo SqlException, tienen un formato especial.
            //----------------------------------------------------------------
            if (this.m_tipoExcepcion == "SqlException")
            {
                oSB.Append(", SQL Error #: " + this.m_sqlErrorNumber);
                oSB.Append(this.m_EspaciadorDefault);
                oSB.Append("Stored Procedure: " + this.m_storedProcedure);
            }

            //----------------------------------------------------------------
            //--- Formatea cada paso de la excepción deacuerdo a la pila.
            //----------------------------------------------------------------
            oSB.Append(this.m_EspaciadorDefault);
            oSB.Append("Descripción: " + this.Message);
            for (int i = this.CallStack.Length - 1; i >= 0; i--)
            {
                oSB.Append(this.m_EspaciadorStack);
                oSB.Append("Frame #" + i.ToString() + ": ");
                oSB.Append(this.CallStack.Frames[i].m_sourceClass + "." + this.CallStack.Frames[i].m_sourceMethod);
            }
            string newMessage = oSB.ToString();

            //----------------------------------------------------------------
            //--- Registra la excepción en un archivo.
            //----------------------------------------------------------------
            if (fileLog)
            {
                string sPath = ConfigurationSettings.AppSettings.Get("ExceptionLogPath");
                if (!System.IO.Directory.Exists(sPath))
                {
                    System.IO.Directory.CreateDirectory(sPath);
                }
                string sPathAndFile = sPath + "\\" + System.DateTime.Now.ToString(this.m_formatoFecha) + "-Exception.log";
                TextWriter oTR = File.AppendText(sPathAndFile);
                oTR.WriteLine(newMessage + oTR.NewLine);
                oTR.Close();
            }

            //----------------------------------------------------------------
            //--- Registra la excepción en EventLog de Windows.
            //----------------------------------------------------------------
            if (eventLog)
            {
                string sEventLogMessage = newMessage.Replace(this.m_EspaciadorDefault, "\n");
                sEventLogMessage = sEventLogMessage.Replace(this.m_EspaciadorStack, "\n");
                EventLog.WriteEntry("Exception", sEventLogMessage, System.Diagnostics.EventLogEntryType.Error);
            }

            return newMessage;
        }


        /// <summary>
        /// Este método permite, de manera opcional, registrar el mensaje de la excepción 
        /// en un archivo o en el EventLog de Windows.
        /// Adicionalmente, formatea el mensaje para desplegarlo en un control tipo 
        /// Label perteneciente a una forma Web.
        /// </summary>
        //public void Log( bool fileLog, bool eventLog, System.Web.UI.WebControls.Label lblMessage)
        //{
        //    //----------------------------------------------------------------
        //    //--- Registra el mensaje en un log.
        //    //----------------------------------------------------------------
        //    string newMessage = Log( fileLog, eventLog);

        //    //----------------------------------------------------------------
        //    //--- Formatea el mensaje para poder mostrarlo en un control tipo Label.
        //    //----------------------------------------------------------------
        //    //string spacer = "<br>" + RepeatString("&nbsp;", 20);
        //    newMessage = newMessage.Replace(this.m_EspaciadorStack, this.m_EspaciadorHTMLStack);
        //    newMessage = newMessage.Replace(this.m_EspaciadorDefault, this.m_EspaciadorHTMLDefault);

        //    //----------------------------------------------------------------
        //    //--- Display it to the Label
        //    //----------------------------------------------------------------
        //    lblMessage.Text = newMessage;
        //}


        #endregion MsgException - Métodos Públicos

        #region MsgException - Métodos Privados

        /// <summary>
        /// Set the spacer properties.
        /// </summary>
        private void InitializeSettings()
        {
            this.m_EspaciadorHTMLDefault = "<br>";
            this.m_EspaciadorHTMLStack = "<br>" + RepeatString("&nbsp;", 20);
        }

        /// <summary>
        /// Repeat a string n number of times.
        /// </summary>
        private string RepeatString(string text, int repeticiones)
        {
            string newText = "";
            for (int i = 1; i <= repeticiones; i++)
            {
                newText += text;
            }
            return newText;
        }


        #endregion MsgException - Métodos Privados
    }
}
