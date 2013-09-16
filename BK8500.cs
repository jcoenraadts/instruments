using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;


namespace Instruments
{
    public class BK8500
    {
        public bool deviceConnected = false;
        private const int packetLength = 26;

        //comms statistics
        public int stats_msgsSent = 0;
        public int stats_expectedReplies = 0;
        public int stats_receivedReplies = 0;
        public int stats_commsRetries = 0;
        public int stats_checkSumFailures = 0;
        public int stats_falseDataRxTriggers = 0;

        public int measureDelay = 10; //ms

        private SerialPort P;
        private EventWaitHandle dataWaitHandle;
        private int dataTimeout = 2000; //ms - comms fail if the instrument does not reply in this time
        private int commsRetries = 3;
        private byte address;
        private const byte STARTBYTE = 0xAA;

        public instrumentStatus status;
        public bool instrumentError { get; set; }

        private enum command : byte
        {
            statusPacket = 0x12,
            remoteOperation = 0x20,
            loadOnOff = 0x21,
            maxVoltSet = 0x22,
            maxVoltRead = 0x23,
            maxCurrentSet = 0x24,
            maxCurrentRead = 0x25,
            maxPowerSet = 0x26,
            maxPowerRead = 0x27,
            regModeSet = 0x28,
            regModeRead = 0x29,
            CC_currentSet = 0x2A,
            CC_currentRead = 0x2B,
            CV_voltSet = 0x2C,
            CV_voltRead = 0x2D,
            CP_powerSet = 0x2E,
            CP_powerRead = 0x2F,
            CR_resistanceSet = 0x30,
            CR_resistanceRead = 0x31,
            //no transients required
            commAddrSet = 0x54,
            localControlSet = 0x55,
            remoteSenseSet = 0x56,
            remoteSenseRead = 0x57,
            valueRead = 0x5F,
            productInfo = 0x6A
        }

        public enum mode : byte
        {
            constantCurrent = 0x00, constantVoltage = 0x01, constantPower = 0x2, constantResistance = 0x03
        }

        private enum statusByte : byte
        {
            checksumIncorrect = 0x90, parameterIncorrect = 0xA0, unrecognisedCommand = 0xB0, invalidCommand = 0xC0, commandOK = 0x80
        }

        public BK8500(string comPortname, byte address)
        {
            instrumentError = false;

            //Setup serial port
            P = new SerialPort(comPortname, 38400);
            P.DataBits = 8;
            P.StopBits = StopBits.One;
            P.Parity = Parity.None;
            P.RtsEnable = true;
            P.ReadTimeout = 200;

            //Setup event for received data
            P.ReceivedBytesThreshold = packetLength;  //all packets for this device are 26 bytes in length
            P.DataReceived += new SerialDataReceivedEventHandler(P_DataReceived);
            dataWaitHandle = new EventWaitHandle(false, EventResetMode.ManualReset);  //this syncs the received data event handler to the main thread

            P.Open();
            P.DiscardInBuffer();
            this.address = address;

            //enter local mode
            this.remoteOperation = true;
            this.loadON = false;
            this.deviceConnected = true;
        }

        #region configuration

        public string deviceInfo
        {
            get
            {
                byte[] b = msgSendGetReply(command.productInfo);
                
                byte[] model = new byte[5];
                Array.Copy(b, 0, model, 0, 5);
                string sModel = System.Text.Encoding.ASCII.GetString(model);
                sModel.Trim();

                byte[] SN = new byte[10];
                Array.Copy(b, 10, SN, 0, SN.Length);
                string sSerial = System.Text.Encoding.ASCII.GetString(SN);
                sSerial.Trim();

                return sModel + " Firmware version: " + b[8].ToString() + "." + b[9].ToString() + " Serial Number: " +sSerial;
            }
        }

        public bool remoteOperation
        {
            set
            {
                byte[] b = {Convert.ToByte(value)};
                msgSend(command.remoteOperation, b);    //1 for remote operation, 0 for local
            }
        }

        public bool loadON
        {
            set
            {
                byte[] b = { Convert.ToByte(value) };
                msgSend(command.loadOnOff, b);    //1 for load on, 0 for laod off
            }
        }

        public double maxVoltage_V
        {
            set
            {
                UInt32 mV = (UInt32)(value * 1000);
                byte[] b = BitConverter.GetBytes(mV);   //little endian array
                msgSend(command.maxVoltSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.maxVoltRead);
                UInt32 mV = BitConverter.ToUInt32(b, 0);
                return ((double)mV / 1000);
            }
        }

        public double maxCurrent_A
        {
            set
            {
                UInt32 mA_10 = (UInt32)(value * 10000);     //current is represented by 0.1mA increments
                byte[] b = BitConverter.GetBytes(mA_10);   //little endian array
                msgSend(command.maxCurrentSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.maxCurrentRead); //current is represented by 0.1mA increments
                UInt32 mA_10 = BitConverter.ToUInt32(b, 0);
                return ((double)mA_10 / 10000);
            }
        }

       
        public double maxPower_W
        {
            set
            {
                UInt32 mW = (UInt32)(value * 1000);
                byte[] b = BitConverter.GetBytes(mW);   //little endian array
                msgSend(command.maxPowerSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.maxPowerRead);
                UInt32 mW = BitConverter.ToUInt32(b, 0);
                return ((double)mW / 1000);
            }
        }
 
        private mode _mode;
        public mode setMode
        {
            set
            {
                byte[] b = { (byte)value };
                msgSend(command.regModeSet, b);
                _mode = value;
            }
            get
            {
                return _mode;
            }
        }

        public double currentSetpoint_A
        {
            set
            {
                UInt32 val = (UInt32)(value * 10000);   //TODO enter min settable
                byte[] b = BitConverter.GetBytes(val);   //little endian array
                msgSend(command.CC_currentSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.CC_currentRead);
                UInt32 val = BitConverter.ToUInt32(b, 0);
                return ((double)val / 10000);
            }
        }

        public double voltageSetpoint_V
        {
            set
            {
                UInt32 val = (UInt32)(value * 1000);
                if (val < 100)
                    val = 100;                           //minimum settable voltage
                byte[] b = BitConverter.GetBytes(val);   //little endian array
                msgSend(command.CV_voltSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.CV_voltRead);
                UInt32 val = BitConverter.ToUInt32(b, 0);
                return ((double)val / 1000);
            }
        }

        public double powerSetpoint_W
        {
            set
            {
                UInt32 val = (UInt32)(value * 1000);        //TODO enter min settable 
                byte[] b = BitConverter.GetBytes(val);   //little endian array
                msgSend(command.CP_powerSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.CP_powerRead);
                UInt32 val = BitConverter.ToUInt32(b, 0);
                return ((double)val / 1000);
            }
        }

        public double impedanceSetpoint_Ohms
        {
            set
            {
                UInt32 val = (UInt32)(value * 1000);        //TODO enter min settable
                byte[] b = BitConverter.GetBytes(val);   //little endian array
                msgSend(command.CR_resistanceSet, b);
            }
            get
            {
                byte[] b = msgSendGetReply(command.CR_resistanceRead);
                UInt32 val = BitConverter.ToUInt32(b, 0);
                return ((double)val / 1000);
            }
        }

        public bool remoteSensing
        {
            set
            {
                byte[] b = { Convert.ToByte(value) };
                msgSend(command.remoteSenseSet, b);    //1 for load on, 0 for laod off
            }
            get
            {
                byte[] b = msgSendGetReply(command.remoteSenseRead);
                return Convert.ToBoolean(b[0]);
            }
        }
        #endregion

        #region readPowerValues

        public void readVIPs(out double voltage, out double current, out double power) 
        {
            //reply: U32 millivolts, U32 mA/10, U32 mW, U8 status register, U16 demand state register
            byte[] b = msgSendGetReply(command.valueRead);

            voltage = (double)BitConverter.ToUInt32(b, 0) / 1000;
            current = (double)BitConverter.ToUInt32(b, 4) / 10000;
            power = (double)BitConverter.ToUInt32(b, 8) / 1000;

            UInt16 dState = BitConverter.ToUInt16(b, 12);
            status = new instrumentStatus(dState);  //update global instrument status
            instrumentError = status.error;
        }

        public double displayVoltage
        {
            get
            {
                double v, i, p;
                readVIPs(out v, out i, out p);
                return v;
            }

        }

        public double displayCurrent
        {
            get
            {
                double v, i, p;
                readVIPs(out v, out i, out p);
                return i;
            }

        }

        public double readVoltage()
        {
            double v, i, p;
            readVIPs(out v, out i, out p);
            return v;
        }

        public double readCurrent()
        {
            double v, i, p;
            readVIPs(out v, out i, out p);
            return i;
        }

        public double readPower()
        {
            double v, i, p;
            readVIPs(out v, out i, out p);
            return p;
        }
        #endregion

        #region communication

        /// <summary>
        /// Sends 24 byte packet to the instrument
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="data"></param>
        private void msgSend(command cmd, byte[] data)
        {
            msgSendGetReply(cmd, data);
        }

        /// <summary>
        /// Sends a packet to the instrument and expects a reply
        /// </summary>
        /// <param name="cmd">Enumerated type of instrument command</param>
        /// <param name="data">Data packet to send to the device</param>
        /// <returns>Data packet returned by the device</returns>
        private byte[] msgSendGetReply(command cmd, byte[] data)
        {
            byte[] b = new byte[packetLength];

            //packet: startbyte, addr, command, data[..], checksum
            b[0] = STARTBYTE;
            b[1] = address;
            b[2] = (byte)cmd;
            Array.Copy(data, 0, b, 3, data.Length);
            b[25] = (byte)(computeChecksum(b));    //checksum is the sum, modulo 256

            return serialTransaction(b);  //just returns the datapacket 
        }

        private byte[] serialTransaction(byte[] inData)
        {
            for (int i = 0; i < commsRetries; i++)      //communication retry loop
            {
                dataWaitHandle.Reset();               //reset the wait event

                P.DiscardInBuffer();
                P.Write(inData, 0, inData.Length);    //clear the buffer and write to the instrument
                this.stats_msgsSent++;

                this.stats_expectedReplies++;
                if (dataWaitHandle.WaitOne(dataTimeout))
                {
                    this.stats_receivedReplies++;

                    //byte[] msg = new byte[P.BytesToRead];   //double buffering means that this value is not always current
                    byte[] msg = new byte[packetLength];                  //number hard coded to get around problem of bytestoread reporting incorrectly
                    for (int count = 0; count < msg.Length; count++)
                        msg[count] = (byte)P.ReadByte();

                    //find message in bytes, read extra bytes if required
                    int index = Array.IndexOf(msg, STARTBYTE);
                    Array.Resize(ref msg, msg.Length + index);
                    for (int j = 0; j < index; j++)
                        msg[j + packetLength] = (byte)P.ReadByte();

                    //set msg array to be just the 26 byte message
                    if (msg.Length > packetLength)
                    {
                        byte[] tempArray = new byte[packetLength];
                        Array.Copy(msg, index, tempArray, 0, tempArray.Length);
                        msg = tempArray;
                    }
                    
                    byte[] dataPacket = new byte[20];
                    Array.Copy(msg, 3, dataPacket, 0, dataPacket.Length);

                    //check checksum
                    byte[] checkPkt = new byte[msg.Length - 1];                 //always last byte
                    Array.Copy(msg, 0, checkPkt, 0, checkPkt.Length);
                    if (computeChecksum(checkPkt) != msg[msg.Length - 1])      //the checksum should always add with the rest of the bytes to zero
                    {
                        stats_checkSumFailures++;
                        continue;               //if the checksum fails, continue to the next iteration of the loop
                    }

                    //if the packet is a status packet (if the requested operation is a GET) the command is 0x12, check the status
                    if (msg[2] == (byte)command.statusPacket)
                    {
                        statusByte s = (statusByte)msg[3];
                        switch (s)
                        {
                            case statusByte.checksumIncorrect:  //if the checksum was incorrect, retry (comms error on lines)
                                continue;
                            case statusByte.commandOK:
                                break;
                            default:
                                throw new Exception("Communication failed, instrument returned status: " + s.ToString());
                        }
                    }
                    
                    return dataPacket;                         //if the data was received and the checksum was OK this loop returns
                }
                System.Threading.Thread.Sleep(100);
            }
            throw new Exception("Communication with the BK8500 has failed, please check all connections and restart the program");
        }

        /// <summary>
        /// Adds together all bytes in an array, for an incoming message this should equal zero if message is good.
        /// </summary>
        /// <param name="data">Data array to be summed</param>
        /// <returns>8bit sum of bytes in array</returns>
        private byte computeChecksum(byte[] data)
        {
            byte sum = 0;
            foreach (byte b in data)
                sum += b;
            return sum;
        }


        //overload for messages without data
        private void msgSend(command cmd)
        {
            byte[] b = new byte[0];
            msgSend(cmd, b);
        }

        //overload for messages without data
        private byte[] msgSendGetReply(command cmd)
        {
            byte[] b = new byte[0];
            return msgSendGetReply(cmd, b);
        }

        //This event is triggered when 26 bytes have veen received.
        void P_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort P = (SerialPort)sender;
            if (P.BytesToRead < packetLength) //this event seems to fire for no reason
            {
                stats_falseDataRxTriggers++;
                return;
            }

            dataWaitHandle.Set();
        }

        public void close()
        {
            if (P.IsOpen)
                P.Close();
        }

        #endregion

        public struct instrumentStatus
        {
            public instrumentStatus(UInt16 demandState)
            {
                int i = 0;
                this.reversedVoltage = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.overVoltage = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.overPower = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.overTemp = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.remoteSenseDisconnected = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.constantCurrent = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.constantVoltage = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.constantPower = Convert.ToBoolean((demandState >> i++) & 0x0001);
                this.constantResistance = Convert.ToBoolean((demandState >> i++) & 0x0001);

                this.error = (this.reversedVoltage ||
                    this.overVoltage ||
                    this.overPower ||
                    this.overTemp ||
                    this.remoteSenseDisconnected);
            }

            public bool error;
            public bool reversedVoltage;
            public bool overVoltage;
            public bool overPower;
            public bool overTemp;
            public bool remoteSenseDisconnected;
            public bool constantCurrent;
            public bool constantVoltage;
            public bool constantPower;
            public bool constantResistance;
        }

        #region measurements

        /// <summary>
        /// Returns the panel curve, by setting the voltage and measureing cuttring in 100 steps down from Voc
        /// </summary>
        /// <param name="voltages">100 voltage steps</param>
        /// <param name="currents">Currents measured at each </param>
        public void measurePanelCurve(int numSteps, out double[] voltages, out double[] currents, out double[] powers)
        {
            //measure Voc
            this.loadON = false;
            double Voc = this.displayVoltage;

            //divide Voc by 100, and sweep through values, measuring current at each point. 
            this.setMode = BK8500.mode.constantVoltage; //set CV mode
            currents = new double[numSteps];
            voltages = new double[numSteps];
            powers = new double[numSteps];
            for (int i = numSteps; i > 0; i--)
            {
                voltages[i - 1] = (Voc * i) / numSteps;
                this.voltageSetpoint_V = voltages[i - 1];       //set voltage setpoint
                if (i == numSteps) this.loadON = true;   //first loop, turn this on
                currents[i - 1] = this.displayCurrent;
                powers[i - 1] = currents[i - 1] * voltages[i - 1];      //calc power
                System.Threading.Thread.Sleep(measureDelay);
            }

            //set load to off
            this.loadON = false;
        }

        #endregion
    }
}
