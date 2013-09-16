using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Ports;
using System.Threading;


namespace Instruments
{
    public class ICS_interface
    {
        private SerialPort P;
        private ICS7017R module1;
        private ICS7017R module2;

        public ICS_interface(string comPortname, string[] addresses)
        {
            //Setup serial port
            P = new SerialPort(comPortname, 9600);
            P.DataBits = 8;
            P.StopBits = StopBits.One;
            P.Parity = Parity.None;
            P.RtsEnable = true;
            P.ReadTimeout = 2000;
            P.NewLine = "\r";
            
            //Currently this class creates two devices on the RS485 bus, this should be adjusted to your needs
            module1 = new ICS7017R(addresses[0], this);
            module2 = new ICS7017R(addresses[1], this);

        }

        /// <summary>
        /// Reads the voltages from both of the I-7017s. Returns true if the data was returned OK
        /// </summary>
        /// <param name="voltages">0-based array of the channel voltages (16 long) connected to this device</param>
        /// <returns></returns>
        public bool readAllVoltages(out double[] voltages)
        {
            bool dataOK = true;

            double[] v1, v2;

            dataOK &= module1.readAllVoltages(out v1);
            dataOK &= module2.readAllVoltages(out v2);
            voltages = new double[v1.Length + v2.Length];

            if (!dataOK)
                return false;
            
            Array.Copy(v1, 0, voltages, 0, v1.Length);
            Array.Copy(v2, 0, voltages, v1.Length, v2.Length);
            return dataOK;
        }

        private bool serialTransaction(string txData, out string rxData)
        {
            //test HACK
            //rxData = ">+025.12+020.45+012.78+018.97+003.24+015.35+018.97+003.24FF";
            //return true;

            rxData = "";

            this.open();
            //P.DiscardInBuffer();
            P.WriteLine(txData);

            bool msgOK = false;
            try
            {
                rxData = P.ReadLine();
                msgOK = true;
            }
            catch (TimeoutException)
            {
                msgOK = false;
            }
            finally
            {
                this.close();
            }
            return msgOK;
        }

        public void close()
        {
            if (P.IsOpen)
                P.Close();
        }

        public void open()
        {
            if (!P.IsOpen)
                P.Open();
        }

        public class ICS7017R
        {
            public bool deviceConnected = false;

            private int commsRetries = 3;
            private string address;
            private ICS_interface iface;

            /// <summary>
            /// Creates an object representing two of these devices
            /// </summary>
            /// <param name="address">Should be of the form "AA"</param>
            public ICS7017R(string address, ICS_interface iface)
            {
                this.address = address;
                this.iface = iface;
            }

            #region readPowerValues

            /// <summary>
            /// Reads the voltages from the I-7017. Returns true if the data was returned OK
            /// </summary>
            /// <param name="voltages">0-based array of the channel voltages connected to this device</param>
            /// <returns></returns>
            public bool readAllVoltages(out double[] voltages)
            {
                //Syntax: #AA[checksum][CR]
                voltages = new double[8];

                string rxPayload;
                if (!sendMsgGetReply("#" + address, out rxPayload))      //if this fails, false is returned
                    return false;

                //data packet returns data in this format:
                //>+025.12+020.45+012.78+018.97+003.24

                char[] seperators = { '+', '-' };
                string[] values = rxPayload.Split(seperators, StringSplitOptions.RemoveEmptyEntries);

                bool parsedOK = true;
                for (int i = 0; i < voltages.Length; i++)
                    parsedOK &= Double.TryParse(values[i], out voltages[i]);

                return parsedOK;
            }

            #endregion

            #region communication
            /// <summary>
            /// Returns the sum of the bytes in cmd as a string
            /// </summary>
            /// <param name="cmd">Full message, but without the [CR]</param>
            /// <returns></returns>
            private string getChecksum(string cmd)
            {
                //Checksum disabled in the hardware!
                return "";
                
                //get the bytes from the message and add them together
                ASCIIEncoding A = new ASCIIEncoding();
                byte[] cmdBytes = A.GetBytes(cmd);
                byte checkSum = 0;
                foreach (byte b in cmdBytes)
                    checkSum += b;

                return checkSum.ToString("X2");
            }

            private bool sendMsgGetReply(string cmd, out string response)
            {
                response = "";
                cmd += getChecksum(cmd);

                string rx;
                for (int i = 0; i < commsRetries; i++)
                {
                    if (!iface.serialTransaction(cmd, out rx))    //send and get reply
                        continue;

                    //decode the rx message
                    string rxHeader = rx.Substring(0, 1);   //returns the start byte
                    //string rxChecksum = rx.Substring(rx.Length - 2, 2); //last two bytes
                    //string rxPayload = rx.Substring(rxHeader.Length, rx.Length - (rxHeader.Length + rxChecksum.Length));

                    //checksum is not used, uncomment above if it is
                    string rxPayload = rx.Substring(rxHeader.Length, rx.Length - rxHeader.Length);
                    response = rxPayload;
                    return true;

                    //if (rxChecksum == getChecksum(rxHeader + rxPayload))
                    //{
                    //    response = rxPayload;
                    //    return true;
                    //}
                }
                return false;
            }

            #endregion

        }
    }
}
