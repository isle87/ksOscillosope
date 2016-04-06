using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ivi.Visa.Interop;
using System.Runtime.InteropServices;

namespace ksScope
{
    public class Scope
    {
        private FormattedIO488 m_IoObject;
        private ResourceManager m_ResourceManager;
        private string m_strVisaAddress;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="VisaAdress">Adress Oscilloscope</param>
        public Scope(string VisaAdress)
        {
            // Save VISA addres in member variable.
            m_strVisaAddress = VisaAdress;
            // Open the default VISA COM IO object.
            OpenIo();
            // Clear the interface.
            m_IoObject.IO.Clear();
        }

        /// <summary>
        /// The Clear common command clears the status data structures, the
        /// device- defined error queue, and the Request- for- OPC flag.
        /// </summary>
        public void Clear()
        {
            DoCommand("*CLS");
        }

        //TODO

        #region Normal
        public void Close()
        {
            try
            {
                m_IoObject.IO.Close();
            }
            catch { }
            try
            {
                Marshal.ReleaseComObject(m_IoObject);
            }
            catch { }
            try
            {
                Marshal.ReleaseComObject(m_ResourceManager);
            }
            catch { }
        }

        /// <summary>
        /// Send a Command
        /// </summary>
        /// <param name="Command">Command String</param>
        public void DoCommand(string Command)
        {
            // Send the command.
            m_IoObject.WriteString(Command, true);
            // Check for inst errors.
            CheckInstrumentErrors(Command);
        }

        /// <summary>
        /// IEEE Block Command
        /// </summary>
        /// <param name="Command">Command</param>
        /// <param name="DataArray">Data Array</param>
        public void DoCommandIEEEBlock(string Command, byte[] DataArray)
        {
            // Send the command to the device.
            m_IoObject.WriteIEEEBlock(Command, DataArray, true);
            // Check for inst errors.
            CheckInstrumentErrors(Command);
        }

        /// <summary>
        /// Send a Query (IEEE Block)
        /// </summary>
        /// <param name="Query"></param>
        /// <returns></returns>
        public byte[] DoQueryIEEEBlock(string Query)
        {
            // Send the query.
            m_IoObject.WriteString(Query, true);
            // Get the results array.
            System.Threading.Thread.Sleep(2000); // Delay before reading.
            byte[] ResultsArray;
            ResultsArray = (byte[])m_IoObject.ReadIEEEBlock(
            IEEEBinaryType.BinaryType_UI1, false, true);
            // Check for inst errors.
            CheckInstrumentErrors(Query);
            // Return results array.
            return ResultsArray;
        }

        /// <summary>
        /// Send a Query (Number)
        /// </summary>
        /// <param name="Query">Query String</param>
        /// <returns>Answer in Numberformat</returns>
        public double DoQueryNumber(string Query)
        {
            // Send the query.
            m_IoObject.WriteString(Query, true);
            // Get the result number.
            double fResult;
            fResult = (double)m_IoObject.ReadNumber(
            IEEEASCIIType.ASCIIType_R8, true);
            // Check for inst errors.
            CheckInstrumentErrors(Query);
            // Return result number.
            return fResult;
        }

        /// <summary>
        /// Send a Query (Number Array)
        /// </summary>
        /// <param name="Query">Query</param>
        /// <returns>Answer in Numberarray</returns>
        public double[] DoQueryNumbers(string Query)
        {
            // Send the query.
            m_IoObject.WriteString(Query, true);
            // Get the result numbers.
            double[] fResultsArray;
            fResultsArray = (double[])m_IoObject.ReadList(
            IEEEASCIIType.ASCIIType_R8, ",;");
            // Check for inst errors.
            CheckInstrumentErrors(Query);
            // Return result numbers.
            return fResultsArray;
        }

        /// <summary>
        /// Send a Query (String)
        /// </summary>
        /// <param name="Query">Query String</param>
        /// <returns>Answer in Stringformat</returns>
        public string DoQueryString(string Query)
        {
            // Send the query.
            m_IoObject.WriteString(Query, true);
            // Get the result string.
            string strResults;
            strResults = m_IoObject.ReadString();
            // Check for inst errors.
            CheckInstrumentErrors(Query);
            // Return results string.
            return strResults;
        }
        /// <summary>
        /// Set Timeout in Seconds
        /// </summary>
        /// <param name="Seconds">Seconds</param>
        public void SetTimeoutSeconds(int Seconds)
        {
            m_IoObject.IO.Timeout = Seconds * 1000;
        }

        private void CheckInstrumentErrors(string Command)
        {
            // Check for instrument errors.
            string strInstrumentError;
            bool bFirstError = true;
            do // While not "0,No error".
            {
                m_IoObject.WriteString(":SYSTem:ERRor?", true);
                strInstrumentError = m_IoObject.ReadString();
                if (!strInstrumentError.ToString().StartsWith("+0,"))
                {
                    if (bFirstError)
                    {
                        Console.WriteLine("ERROR(s) for command '{0}': ",
                        Command);
                        bFirstError = false;
                    }
                    Console.Write(strInstrumentError);
                }
            } while (!strInstrumentError.ToString().StartsWith("+0,"));
        }

        private void OpenIo()
        {
            m_ResourceManager = new ResourceManager();
            m_IoObject = new FormattedIO488();
            // Open the default VISA COM IO object.
            try
            {
                m_IoObject.IO =
                (IMessage)m_ResourceManager.Open(m_strVisaAddress,
                AccessMode.NO_LOCK, 0, "");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: {0}", e.Message);
            }
        }
        #endregion



    }
}


