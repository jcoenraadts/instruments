using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace Instruments
{
    public class BK2831E
    {
        #region globals
        private SerialPort P;
        private const int numCommsRetries = 3;
        private string[] modes;

        public int stats_commsRetries = 0;
        public int stats_commsAttempts = 0;

        #endregion

        public BK2831E(string comPort)
        {
            P = new SerialPort(comPort);
            P.BaudRate = 38400;
            P.ReadTimeout = 1000;
            P.Open();

            P.DiscardInBuffer();
            P.DiscardOutBuffer();

            int i = 0;
            modes = new string[9];
            modes[i++] = "VOLTAGE:AC";
            modes[i++] = "VOLTAGE:DC";
            modes[i++] = "CURRENT:AC";
            modes[i++] = "CURRENT:DC";
            modes[i++] = "RESISTANCE";
            modes[i++] = "FREQUENCY";
            modes[i++] = "PERIOD";
            modes[i++] = "DIODE";
            modes[i++] = "CONTINUITY";

            reset();


            //test comms with the device
            try
            {
                string instrID = this.instrumentID;
                if (!instrID.Contains("2831E  Multimeter"))
                    throw new Exception("Instrument expected on " + comPort + " is a BK2831E  Multimeter, " + instrID + " was found");
            }
            catch (TimeoutException)
            {
                throw new Exception("No instrument was found on " + comPort + ", A BK2831E was expected");
            }
        }

        #region control

        /// <summary>
        /// Writes a string to the DMM via a serialport object, the string should be reflected by the instrument
        /// </summary>
        /// <param name="instruction">String to write to port</param>
        private void writeSCPI(string instruction)
        {
            int i = 0;
            while (true)
            {
                try
                {
                    stats_commsAttempts++;
                    P.WriteLine(instruction);
                    string read = P.ReadLine();
                    if (read.Contains(instruction))    //if the string read from the instrument does not match the instruction, retry.
                        break;
                    stats_commsRetries++;
                }
                catch (TimeoutException e)
                {
                    if (i == numCommsRetries)
                        throw e;
                    stats_commsRetries++;
                }
                i++;
            }
        }

        /// <summary>
        /// Uses writeSCPI to send a string to the instrument, and waits for a reply. Retries three times
        /// </summary>
        /// <param name="instruction">String to send to the instrument (SCPI)</param>
        /// <returns>The string received from the instrument</returns>
        private string readWriteSCPI(string instruction) 
        {
            int i = 0;
            while(true)
            {
                try
                {
                    writeSCPI(instruction);
                    string s = P.ReadLine();
                    return s;
                }
                catch(TimeoutException e)
                {
                    if (i == numCommsRetries)
                        throw e;
                    stats_commsRetries++;
                }
                i++;
            }
             
        }

        /// <summary>
        /// Closes the serial connection to the instrument
        /// </summary>
        public void close()
        {
            if (P.IsOpen)
                P.Close();
        }

        #endregion

        #region commonCommands

        public string instrumentID
        {
            get
            {
                return readWriteSCPI("*IDN?");
            }
        }

        public void reset()
        {
            writeSCPI("*RST");
            System.Threading.Thread.Sleep(2500);
            P.Close();
            System.Threading.Thread.Sleep(100);
            P.Open();

            P.DiscardInBuffer();
            P.DiscardOutBuffer();

        }

        #endregion

        #region displaySubsystem

        public bool displayEnabled
        {
            set
            {
                string s = ":DISPLAY:ENABLE " + Convert.ToInt16(value).ToString();
                writeSCPI(s);
            }
            get
            {
                string reply = readWriteSCPI(":DISPLAY:ENABLE?");
                if ((reply == "1") || (reply == "ON"))
                    return true;
                else
                    return false;
            }
        }

        #endregion

        #region functionSubsystem

        public enum function
        {
            voltageAC, voltageDC, currentAC, currentDC, resistance, frequency, period, diode, continuity
        }

        public function Function
        {
            set
            {
                string instr = ":FUNCTION " + modes[(int)value];
                writeSCPI(instr);
            }
            get
            {
                string reply = readWriteSCPI(":FUNCTION?");
                for (int i = 0; i < modes.Length; i++)
                    if (reply.Contains(modes[i]))
                        return (function)i;
                
                return (function)(0);   //this line has no effect as the loop should always force a return
            }
        }

        #endregion

        #region voltageSubsystem

        public enum voltageIntegRate { fast = 1, medium = 10, slow = 100 }  //x10 as fast = 0.1 cycles
        public enum voltageRange { _200mV = 2, _2V = 20, _20V = 200, _200V = 2000, _max = 10000 } //x10 as 200mV is 0.2V

        public voltageIntegRate voltageintegRate_AC
        {
            set
            {
                double rate = (double)value / 10;
                string s = "VOLTAGE:AC:NPLCYCLES " + rate.ToString();
                writeSCPI(s);
            }
        }

        public voltageIntegRate voltageintegRate_DC
        {
            set
            {
                double rate = (double)value / 10;
                string s = "VOLTAGE:DC:NPLCYCLES " + rate.ToString();
                writeSCPI(s);
            }
        }

        public voltageRange voltageRange_AC
        {
            set
            {
                double d = (double)value / 10;
                if (value == voltageRange._max) //max is different for AC/DC
                    d = 750;
                string s = "VOLTAGE:AC:RANGE " + d.ToString();
                writeSCPI(s);
            }
        }

        public voltageRange voltageRange_DC
        {
            set
            {
                double d = (double)value / 10;
                string s = "VOLTAGE:DC:RANGE " + d.ToString();
                writeSCPI(s);
            }
        }

        #endregion

        #region measure

        /// <summary>
        /// Returns the value returned by the instrument, regardless of the mode
        /// </summary>
        /// <returns>A double of voltage, current or resistance</returns>
        public double measure()
        {
            string s = readWriteSCPI(":FETCH?");
            return Convert.ToDouble(s);
        }

        #endregion
    }
}
