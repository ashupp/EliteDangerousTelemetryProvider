using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using ILogger = SimFeedback.log.ILogger;
using System.Runtime.InteropServices;

namespace SimFeedback.telemetry
{
    public sealed class TelemetryProvider : AbstractTelemetryProvider
    {
        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int OpenProcess(int dwDesiredAccess, int bInheritHandle, int dwProcessId);

        [DllImport("wow64ext.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern ulong GetBaseAddress64(uint procId);

        [DllImport("wow64ext.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern ulong ReadProcessMemory64(int hProcess, ulong lpBaseAddress, byte[] lpBuffer, int nSize, ref int lpNumberOfBytesRead);

        private bool _isStopped = true;
        private Thread _t;

        private const int ProcessAllAccess = 2035711;
        private string _processName = "EliteDangerous64";

        private int _baseProcess;
        private ulong _baseAddress64;
        private bool _adressesLoaded;

        #region pointers Flying
        private ulong baseElemPoint = 0x041B4688;
        private ulong baseDataSize = 0x8;
        private ulong baseElementSize = 0x368;
        private ulong baseElementOffset = 0x4B8;
        private ulong _heaveOffset = 0x3BC;
        private ulong _pitchOffset = 0x3D8;
        private ulong _rollOffset = 0x3E0;
        private ulong _speedOffset = 0x150;
        private ulong _surgeOffset = 0x3C0;
        private ulong _swayOffset = 0x3B8;
        private ulong _yawOffset = 0x3DC;
        private ulong _surgeAddress;
        private ulong _swayAddress;
        private ulong _heaveAddress;
        private ulong _pitchAddress;
        private ulong _rollAddress;
        private ulong _yawAddress;
        private ulong _speedAddress;
        #endregion

        #region pointers Driving
        private ulong baseElemPointSRV = 0x041AAC08;
        private ulong baseDataSizeSRV = 0x30;
        private ulong baseElementSizeSRV = 0x268;
        private ulong baseElementOffsetSRV = 0x428;
        private ulong baseElementOffsetSRVPitchRollA = 0x4C0;
        private ulong baseElementOffsetSRVPitchRollB = 0x5B0;
        private ulong baseElementOffsetSRVPitchRollC = 0x70;

        private ulong _heaveOffsetSRV = 0x55C;
        private ulong _pitchOffsetSRV = 0x3FC;
        private ulong _rollOffsetSRV = 0x400;
        private ulong _speedOffsetSRV = 0x560;
        private ulong _surgeOffsetSRV = 0x5A0;
        private ulong _swayOffsetSRV = 0x558;
        private ulong _yawOffsetSRV = 0x5C0;
        private ulong _surgeAddressSRV;
        private ulong _swayAddressSRV;
        private ulong _heaveAddressSRV;
        private ulong _pitchAddressSRV;
        private ulong _rollAddressSRV;
        private ulong _yawAddressSRV;
        private ulong _speedAddressSRV;
        #endregion

        public TelemetryProvider()
        {
            Author = "ashupp / ashnet GmbH";
            Version = Assembly.LoadFrom(Assembly.GetExecutingAssembly().Location).GetName().Version.ToString();
            BannerImage = @"img\banner_elitedangerous.png";
            IconImage = @"img\icon_elitedangerous.png";
            TelemetryUpdateFrequency = 60;
            _baseProcess = -1;
        }

        public override string Name => "elitedangerous";

        public override void Init(ILogger logger)
        {
            base.Init(logger);
            Log("Initializing " + Name + "TelemetryProvider");
        }

        public override string[] GetValueList()
        {
            return GetValueListByReflection(typeof(TelemetryData));
        }

        public override void Stop()
        {
            if (_isStopped) return;
            LogDebug("Stopping " + Name + "TelemetryProvider");
            _isStopped = true;
            _baseProcess = -1;
            if (_t != null) _t.Join();
        }

        public override void Start()
        {
            if (_isStopped)
            {
                LogDebug("Starting " + Name + "TelemetryProvider");
                _isStopped = false;
                _t = new Thread(Run);
                _t.Start();
            }
        }

        private void Run()
        {
            TelemetryData lastTelemetryData = new TelemetryData();
            Stopwatch sw = new Stopwatch();
            Stopwatch swAddressHelper = new Stopwatch();
            sw.Start();
            swAddressHelper.Start();

            while (!_isStopped)
            {
                if (sw.ElapsedMilliseconds > 500)
                {
                    IsRunning = false;
                    Thread.Sleep(1000);
                }

                try
                {
                    if (_baseProcess == -1 || _adressesLoaded == false)
                    {
                        IsConnected = false;
                        try
                        {
                            _baseProcess = -1;
                            _adressesLoaded = false;
                            InternalGameStart();

                        }
                        catch (Exception ex)
                        {
                            LogDebug("Error:" + ex.Message);
                        }
                        Thread.Sleep(1000);
                    }
                    else
                    {

                        if (swAddressHelper.ElapsedMilliseconds > 1000)
                        {
                            GetDataAddresses();
                            swAddressHelper.Restart();
                        }


                        TelemetryData telemetryData = new TelemetryData();

                        // Flying
                        var Surge = GetSingle(_surgeAddress);
                        var Sway = GetSingle(_swayAddress);
                        var Heave = GetSingle(_heaveAddress);
                        var Pitch = GetSingle(_pitchAddress);
                        var Roll = GetSingle(_rollAddress);
                        var Yaw = GetSingle(_yawAddress);
                        var Speed = GetSingle(_speedAddress);

                        
                        if (Surge + Sway + Heave + Pitch + Roll + Yaw == 0.0)
                        {
                            // Driving
                            Surge = (float)(GetSingle(_surgeAddressSRV) / 9.80665);
                            Sway = (float)(GetSingle(_swayAddressSRV) / 9.80665);
                            Heave = (float)(GetSingle(_heaveAddressSRV) / 9.80665);
                            Pitch = GetSingle(_pitchAddressSRV);
                            Roll = GetSingle(_rollAddressSRV);
                            Yaw = GetSingle(_yawAddressSRV);
                            Speed = GetSingle(_speedAddressSRV);
                        }


                        telemetryData.Surge = Surge;
                        telemetryData.Sway = Sway;
                        telemetryData.Heave = Heave;
                        telemetryData.Pitch = Pitch;
                        telemetryData.Roll = Roll;
                        telemetryData.Yaw = Yaw;
                        telemetryData.Speed = Speed;

                        IsConnected = true;
                        IsRunning = true;



                        TelemetryEventArgs args = new TelemetryEventArgs(new TelemetryInfoElem(telemetryData, lastTelemetryData));
                        RaiseEvent(OnTelemetryUpdate, args);
                        lastTelemetryData = telemetryData;
                        sw.Restart();

                        
                        Thread.Sleep(SamplePeriod);
                    }

                }
                catch (Exception e)
                {
                    LogError(Name + "TelemetryProvider Exception while processing data", e);
                    IsConnected = false;
                    IsRunning = false;
                    Thread.Sleep(1000);
                }
            }

            IsConnected = false;
            IsRunning = false;
        }

        private void InternalGameStart()
        {
            _adressesLoaded = false;
            try
            {
                _baseProcess = OpenProcess(ProcessAllAccess, 1, Process.GetProcessesByName(_processName)[0].Id);
                _baseAddress64 = GetBaseAddress64((uint)Process.GetProcessesByName(_processName)[0].Id);
                GetDataAddresses();
            }
            catch (Exception ex)
            {
                _adressesLoaded = false;
                LogError("Error in TelemetryProvider: " + ex.Message);
            }
        }

        #region GetAddress helpers 
        private ulong GetAddressHelperSrvPitchRoll(ulong elemOffset)
        {
            var elemBaseAddress = _baseAddress64;
            elemBaseAddress = GetLong(elemBaseAddress + baseElemPointSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseDataSizeSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementSizeSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementOffsetSRVPitchRollA);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementOffsetSRVPitchRollB);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementOffsetSRVPitchRollC);

            elemBaseAddress = elemBaseAddress + elemOffset;
            return elemBaseAddress;
        }

        private ulong GetAddressHelper(ulong elemOffset)
        {
            var elemBaseAddress = _baseAddress64;

            elemBaseAddress = GetLong(elemBaseAddress + baseElemPoint);
            elemBaseAddress = GetLong(elemBaseAddress + baseDataSize);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementSize);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementOffset);

            elemBaseAddress = elemBaseAddress + elemOffset;
            return elemBaseAddress;
        }

        private ulong GetAddressHelperSrv(ulong elemOffset)
        {
            var elemBaseAddress = _baseAddress64;

            elemBaseAddress = GetLong(elemBaseAddress + baseElemPointSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseDataSizeSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementSizeSRV);
            elemBaseAddress = GetLong(elemBaseAddress + baseElementOffsetSRV);

            elemBaseAddress = elemBaseAddress + elemOffset;
            return elemBaseAddress;
        }

        private void GetDataAddresses()
        {
            try
            {
                // Flying
                _surgeAddress = GetAddressHelper(_surgeOffset);
                _swayAddress = GetAddressHelper(_swayOffset);
                _heaveAddress = GetAddressHelper(_heaveOffset);
                _pitchAddress = GetAddressHelper(_pitchOffset);
                _rollAddress = GetAddressHelper(_rollOffset);
                _yawAddress = GetAddressHelper(_yawOffset);
                _speedAddress = GetAddressHelper(_speedOffset);

                // Driving
                _surgeAddressSRV = GetAddressHelperSrv(_surgeOffsetSRV);
                _swayAddressSRV = GetAddressHelperSrv(_swayOffsetSRV);
                _heaveAddressSRV = GetAddressHelperSrv(_heaveOffsetSRV);
                _pitchAddressSRV = GetAddressHelperSrvPitchRoll(_pitchOffsetSRV);
                _rollAddressSRV = GetAddressHelperSrvPitchRoll(_rollOffsetSRV);
                _yawAddressSRV = GetAddressHelperSrv(_yawOffsetSRV);
                _speedAddressSRV = GetAddressHelperSrv(_speedOffsetSRV);

                _adressesLoaded = true;
            }
            catch (Exception ex)
            {
                LogDebug("Error:" + ex.Message);
            }
        }

        private float GetSingle(ulong dwAddress)
        {
            var singleVal = 0.0f;
            try
            {
                var numArray = new byte[4];
                var lpNumberOfBytesRead = 0;
                ReadProcessMemory64(_baseProcess, dwAddress, numArray, 4, ref lpNumberOfBytesRead);
                singleVal = BitConverter.ToSingle(numArray, 0);
            }
            catch (Exception ex)
            {
                LogError("Error Read Single: " + ex.Message);
            }
            return singleVal;
        }

        private ulong GetLong(ulong dwAddress)
        {
            ulong ulongVal = 0;
            try
            {
                var numArray = new byte[8];
                var lpNumberOfBytesRead = 0;
                ReadProcessMemory64(_baseProcess, dwAddress, numArray, 8, ref lpNumberOfBytesRead);
                ulongVal = BitConverter.ToUInt64(numArray, 0);
            }
            catch (Exception ex)
            {
                LogError("Error Read Long: " + ex.Message);
            }
            return ulongVal;
        }
        #endregion
    }
}