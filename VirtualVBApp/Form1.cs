using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace VirtualVBApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private ushort A_ESV_1 = 0;
        private ushort A_ESV_2 = 0;
        private ushort A_ESV_3 = 0;
        private ushort A_ESV_4 = 0;
        private ushort A_ESV_5 = 0;
        private ushort A_ESV_6 = 0;
        private ushort A_ESV_TOTAL = 0;

        private ushort A_OCV = 0;
        private ushort A_CURRENT = 0;

        private ushort B_ESV_1 = 0;
        private ushort B_ESV_2 = 0;
        private ushort B_ESV_3 = 0;
        private ushort B_ESV_4 = 0;
        private ushort B_ESV_5 = 0;
        private ushort B_ESV_6 = 0;
        private ushort B_ESV_TOTAL = 0;

        private ushort B_OCV = 0;
        private ushort B_CURRENT = 0;




        // 创建四个从站实例，分别为地址11、22、33、44
        ModbusRtuSlave slave1;
        ModbusRtuSlave slave2;


        private void Form1_Load(object sender, EventArgs e)
        {
            this.ActiveControl = null;  // 取消所有控件的焦点
            textBox1.DeselectAll();  // 取消所有选中的文本
            textBox1.SelectionLength = 0;  // 不选中文本
            pictureBox2.Visible = true;
            pictureBox3.Visible = true;

            // 获取所有可用的串口名称
            string[] portNames = SerialPort.GetPortNames();

            // 将串口名称添加到 ComboBox 中
            comboBox1.Items.Clear();  // 清除原有的项
            foreach (string port in portNames)
            {
                comboBox1.Items.Add(port);  // 添加串口名称到 ComboBox
            }

            // 如果有可用串口，默认选择第一个串口
            if (comboBox1.Items.Count > 0)
            {
                comboBox1.SelectedIndex = 0;
            }

            A_ESV_1 = (ushort)trackBar1.Value;
            A_ESV_2 = (ushort)trackBar1.Value;
            A_ESV_3 = (ushort)trackBar1.Value;
            A_ESV_4 = (ushort)trackBar1.Value;
            A_ESV_5 = (ushort)trackBar1.Value;
            A_ESV_6 = (ushort)trackBar1.Value;
            A_ESV_TOTAL = (ushort)(((ushort)trackBar1.Value) * 3);

            B_ESV_1 = (ushort)trackBar2.Value;
            B_ESV_2 = (ushort)trackBar2.Value;
            B_ESV_3 = (ushort)trackBar2.Value;
            B_ESV_4 = (ushort)trackBar2.Value;
            B_ESV_5 = (ushort)trackBar2.Value;
            B_ESV_6 = (ushort)trackBar2.Value;
            B_ESV_TOTAL = (ushort)(((ushort)trackBar2.Value) * 3);

            textBox1.Text = ((float)A_ESV_1 / 10.0).ToString("F1") + "V";
            textBox2.Text = ((float)A_ESV_2 / 10.0).ToString("F1") + "V";
            textBox4.Text = ((float)A_ESV_3 / 10.0).ToString("F1") + "V";
            textBox3.Text = ((float)A_ESV_4 / 10.0).ToString("F1") + "V";
            textBox6.Text = ((float)A_ESV_5 / 10.0).ToString("F1") + "V";
            textBox5.Text = ((float)A_ESV_6 / 10.0).ToString("F1") + "V";
            textBox9.Text = ((float)A_ESV_TOTAL / 10.0).ToString("F1") + "V";


            textBox18.Text = ((float)B_ESV_1 / 10.0).ToString("F1") + "V";
            textBox17.Text = ((float)B_ESV_2 / 10.0).ToString("F1") + "V";
            textBox16.Text = ((float)B_ESV_3 / 10.0).ToString("F1") + "V";
            textBox15.Text = ((float)B_ESV_4 / 10.0).ToString("F1") + "V";
            textBox14.Text = ((float)B_ESV_5 / 10.0).ToString("F1") + "V";
            textBox13.Text = ((float)B_ESV_6 / 10.0).ToString("F1") + "V";
            textBox10.Text = ((float)B_ESV_TOTAL / 10.0).ToString("F1") + "V";


            A_OCV = (ushort)trackBar3.Value;
            B_OCV = (ushort)trackBar6.Value;

            textBox7.Text = ((float)A_OCV / 10000.0).ToString("F4") + "V";
            textBox12.Text = ((float)B_OCV / 10000.0).ToString("F4") + "V";



            A_CURRENT = (ushort)trackBar4.Value;
            B_CURRENT = (ushort)trackBar5.Value;

            textBox8.Text = ((float)A_CURRENT / 10.0).ToString("F1") + "A";
            textBox11.Text = ((float)B_CURRENT / 10.0).ToString("F1") + "A";
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            A_ESV_1 = (ushort)trackBar1.Value;
            A_ESV_2 = (ushort)trackBar1.Value;
            A_ESV_3 = (ushort)trackBar1.Value;
            A_ESV_4 = (ushort)trackBar1.Value;
            A_ESV_5 = (ushort)trackBar1.Value;
            A_ESV_6 = (ushort)trackBar1.Value;
            A_ESV_TOTAL = (ushort)(((ushort)trackBar1.Value) * 3);

            textBox1.Text = ((float)A_ESV_1 / 10.0).ToString("F1") + "V";
            textBox2.Text = ((float)A_ESV_2 / 10.0).ToString("F1") + "V";
            textBox4.Text = ((float)A_ESV_3 / 10.0).ToString("F1") + "V";
            textBox3.Text = ((float)A_ESV_4 / 10.0).ToString("F1") + "V";
            textBox6.Text = ((float)A_ESV_5 / 10.0).ToString("F1") + "V";
            textBox5.Text = ((float)A_ESV_6 / 10.0).ToString("F1") + "V";
            textBox9.Text = ((float)A_ESV_TOTAL / 10.0).ToString("F1") + "V";
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (ushort.TryParse(textBox7.Text, out A_OCV))
            {
                Console.WriteLine("A OCV ->" + A_OCV.ToString());
            }
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            B_ESV_1 = (ushort)trackBar2.Value;
            B_ESV_2 = (ushort)trackBar2.Value;
            B_ESV_3 = (ushort)trackBar2.Value;
            B_ESV_4 = (ushort)trackBar2.Value;
            B_ESV_5 = (ushort)trackBar2.Value;
            B_ESV_6 = (ushort)trackBar2.Value;
            B_ESV_TOTAL = (ushort)(((ushort)trackBar2.Value) * 3);

            textBox18.Text = ((float)B_ESV_1 / 10.0).ToString("F1") + "V";
            textBox17.Text = ((float)B_ESV_2 / 10.0).ToString("F1") + "V";
            textBox16.Text = ((float)B_ESV_3 / 10.0).ToString("F1") + "V";
            textBox15.Text = ((float)B_ESV_4 / 10.0).ToString("F1") + "V";
            textBox14.Text = ((float)B_ESV_5 / 10.0).ToString("F1") + "V";
            textBox13.Text = ((float)B_ESV_6 / 10.0).ToString("F1") + "V";
            textBox10.Text = ((float)B_ESV_TOTAL / 10.0).ToString("F1") + "V";
        }


        private void Update_RTU_Regs()
        {
            slave1.SetHoldingRegister(36, A_ESV_TOTAL);
            slave1.SetHoldingRegister(37, A_ESV_1);
            slave1.SetHoldingRegister(38, A_ESV_2);
            slave1.SetHoldingRegister(39, A_ESV_3);
            slave1.SetHoldingRegister(40, A_ESV_4);
            slave1.SetHoldingRegister(41, A_ESV_5);
            slave1.SetHoldingRegister(42, A_ESV_6);
            slave1.SetHoldingRegister(43, 0);
            slave1.SetHoldingRegister(44, 0);
            slave1.SetHoldingRegister(45, 0);
            slave1.SetHoldingRegister(46, A_OCV);
            slave1.SetHoldingRegister(47, A_CURRENT);


            slave2.SetHoldingRegister(36, B_ESV_TOTAL);
            slave2.SetHoldingRegister(37, B_ESV_1);
            slave2.SetHoldingRegister(38, B_ESV_2);
            slave2.SetHoldingRegister(39, B_ESV_3);
            slave2.SetHoldingRegister(40, B_ESV_4);
            slave2.SetHoldingRegister(41, B_ESV_5);
            slave2.SetHoldingRegister(42, B_ESV_6);
            slave2.SetHoldingRegister(43, 0);
            slave2.SetHoldingRegister(44, 0);
            slave2.SetHoldingRegister(45, 0);
            slave2.SetHoldingRegister(46, B_OCV);
            slave2.SetHoldingRegister(47, B_CURRENT);
        }

        private Random random = new Random();

        private ushort GenerateRandomNumber()
        {
            // 生成一个范围在 0 到 10 之间的随机数
            int randomNumber = random.Next(0, 11);

            // 通过生成的数值来表示范围 -5 到 5
            // 如果 randomNumber 小于 5，则为负数，否则为正数
            int finalNumber = randomNumber - 5;  // -5 到 5 之间的值

            // 如果你需要返回的是 ushort 数字，建议用 Math.Abs 取绝对值（即转换为正数）
            return (ushort)Math.Abs(finalNumber);  // 返回正数
        }

        // 通用的方法来更新值
        private void UpdateESV(CheckBox checkBox1, CheckBox checkBox2, ref ushort esvVariable)
        {
            if (checkBox1.Checked || checkBox2.Checked)
            {
                esvVariable = checkBox1.Checked ? (ushort)1000 : (ushort)0;
            }
        }

        // 确保变量的值不小于 10
        private void EnsureMinimumValue(ref ushort value)
        {
            if (value < 10)
            {
                value = 10;  // 如果值小于 10，设置为 10
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {

            Console.WriteLine("CLICK");

            A_ESV_1 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            A_ESV_2 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            A_ESV_3 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            A_ESV_4 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            A_ESV_5 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            A_ESV_6 = (ushort)(trackBar1.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort

            B_ESV_1 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_ESV_2 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_ESV_3 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_ESV_4 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_ESV_5 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_ESV_6 = (ushort)(trackBar2.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort

            // 确保每个 A_ESV_变量不小于 10
            EnsureMinimumValue(ref A_ESV_1);
            EnsureMinimumValue(ref A_ESV_2);
            EnsureMinimumValue(ref A_ESV_3);
            EnsureMinimumValue(ref A_ESV_4);
            EnsureMinimumValue(ref A_ESV_5);
            EnsureMinimumValue(ref A_ESV_6);


            // 确保每个 A_ESV_变量不小于 10
            EnsureMinimumValue(ref B_ESV_1);
            EnsureMinimumValue(ref B_ESV_2);
            EnsureMinimumValue(ref B_ESV_3);
            EnsureMinimumValue(ref B_ESV_4);
            EnsureMinimumValue(ref B_ESV_5);
            EnsureMinimumValue(ref B_ESV_6);


            UpdateESV(checkBox2, checkBox3, ref A_ESV_1);
            UpdateESV(checkBox4, checkBox1, ref A_ESV_2);
            UpdateESV(checkBox6, checkBox5, ref A_ESV_3);
            UpdateESV(checkBox8, checkBox7, ref A_ESV_4);
            UpdateESV(checkBox10, checkBox9, ref A_ESV_5);
            UpdateESV(checkBox12, checkBox11, ref A_ESV_6);

            UpdateESV(checkBox36, checkBox34, ref B_ESV_1);
            UpdateESV(checkBox35, checkBox33, ref B_ESV_2);
            UpdateESV(checkBox32, checkBox31, ref B_ESV_3);
            UpdateESV(checkBox30, checkBox29, ref B_ESV_4);
            UpdateESV(checkBox28, checkBox27, ref B_ESV_5);
            UpdateESV(checkBox26, checkBox25, ref B_ESV_6);


            A_ESV_TOTAL = Math.Max((ushort)(A_ESV_1 + A_ESV_2 + A_ESV_3), (ushort)(A_ESV_4 + A_ESV_5 + A_ESV_6));
            B_ESV_TOTAL = Math.Max((ushort)(B_ESV_1 + B_ESV_2 + B_ESV_3), (ushort)(B_ESV_4 + B_ESV_5 + B_ESV_6));

            textBox1.Text = ((float)A_ESV_1 / 10.0).ToString("F1") + "V";
            textBox2.Text = ((float)A_ESV_2 / 10.0).ToString("F1") + "V";
            textBox4.Text = ((float)A_ESV_3 / 10.0).ToString("F1") + "V";
            textBox3.Text = ((float)A_ESV_4 / 10.0).ToString("F1") + "V";
            textBox6.Text = ((float)A_ESV_5 / 10.0).ToString("F1") + "V";
            textBox5.Text = ((float)A_ESV_6 / 10.0).ToString("F1") + "V";
            textBox9.Text = ((float)A_ESV_TOTAL / 10.0).ToString("F1") + "V";

            textBox18.Text = ((float)B_ESV_1 / 10.0).ToString("F1") + "V";
            textBox17.Text = ((float)B_ESV_2 / 10.0).ToString("F1") + "V";
            textBox16.Text = ((float)B_ESV_3 / 10.0).ToString("F1") + "V";
            textBox15.Text = ((float)B_ESV_4 / 10.0).ToString("F1") + "V";
            textBox14.Text = ((float)B_ESV_5 / 10.0).ToString("F1") + "V";
            textBox13.Text = ((float)B_ESV_6 / 10.0).ToString("F1") + "V";
            textBox10.Text = ((float)B_ESV_TOTAL / 10.0).ToString("F1") + "V";

            A_CURRENT = (ushort)(trackBar5.Value + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            B_CURRENT = (ushort)(A_CURRENT + (ushort)GenerateRandomNumber());  // Cast random number to ushort
            //
            textBox8.Text = ((float)A_CURRENT / 10.0).ToString("F1") + "A";
            textBox11.Text = ((float)B_CURRENT / 10.0).ToString("F1") + "A";
            //
            A_OCV = (ushort)trackBar3.Value;
            B_OCV = (ushort)trackBar6.Value;
            //
            Update_RTU_Regs();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer2.Enabled = true;

            button1.ForeColor = Color.Green;

            string selectedPort = comboBox1.SelectedItem?.ToString();
            // 设置串口配置
            ModbusRtuSlave.SetSerialPortSettings(selectedPort, 9600, Parity.None, 8, StopBits.One);
            Thread.Sleep(100);

            byte slave_id1 = 1;

            if (byte.TryParse(textBox19.Text, out slave_id1))
            {
                // 如果转换成功，slave_id1 将被赋予 textBox19.Text 转换后的值
            }

            byte slave_id2 = 2;

            if (byte.TryParse(textBox20.Text, out slave_id2))
            {
                // 如果转换成功，slave_id2 将被赋予 textBox19.Text 转换后的值
            }


            // 创建四个从站实例，分别为地址11、22、33、44
            slave1 = new ModbusRtuSlave(slave_id1);
            slave2 = new ModbusRtuSlave(slave_id2);
            //
            Update_RTU_Regs();
            //
            Thread.Sleep(50);

            // 启动共享的 Modbus 线程
            ModbusRtuSlave.Start();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (ushort.TryParse(textBox12.Text, out B_OCV))
            {
                Console.WriteLine("B OCV ->" + B_OCV.ToString());
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            A_OCV = (ushort)trackBar3.Value;
            B_OCV = (ushort)(A_OCV + (ushort)GenerateRandomNumber());
            trackBar6.Value = trackBar3.Value;
            textBox7.Text = ((float)A_OCV / 10000.0).ToString("F4");
            textBox12.Text = ((float)B_OCV / 10000.0).ToString("F4");
        }

        private void trackBar6_Scroll(object sender, EventArgs e)
        {
            B_OCV = (ushort)trackBar6.Value;
            A_OCV = (ushort)(B_OCV + (ushort)GenerateRandomNumber());
            trackBar3.Value = trackBar6.Value;
            textBox7.Text = ((float)A_OCV / 10000.0).ToString("F4");
            textBox12.Text = ((float)B_OCV / 10000.0).ToString("F4");
        }

        private void trackBar4_Scroll(object sender, EventArgs e)
        {
            A_CURRENT = (ushort)trackBar4.Value;
            B_CURRENT = (ushort)A_CURRENT;
            trackBar5.Value = trackBar4.Value;

            textBox8.Text = ((float)A_CURRENT / 10.0).ToString("F1") + "A";
            textBox11.Text = ((float)B_CURRENT / 10.0).ToString("F1") + "A";
        }

        private void trackBar5_Scroll(object sender, EventArgs e)
        {
            trackBar4.Value = trackBar5.Value;
            A_CURRENT = (ushort)trackBar5.Value;
            B_CURRENT = (ushort)A_CURRENT;

            textBox8.Text = ((float)A_CURRENT / 10.0).ToString("F1") + "A";
            textBox11.Text = ((float)B_CURRENT / 10.0).ToString("F1") + "A";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //退出程序
            Application.Exit();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (slave1.GetCommState())
            {
                pictureBox2.Visible = true;
            }
            else
            {
                pictureBox2.Visible = false;
            }

            if (slave2.GetCommState())
            {
                pictureBox3.Visible = true;
            }
            else
            {
                pictureBox3.Visible = false;
            }
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }
    }
}
public class ModbusRtuSlave
{
    private static bool isPortOpen = false;  // 是否已打开串口
    private static SerialPort serialPort;    // 串口实例
    private static Dictionary<byte, ModbusRtuSlave> slaveInstances = new Dictionary<byte, ModbusRtuSlave>();  // 存储所有从站实例

    private byte slaveAddress;
    private ushort[] holdingRegisters;

    private float currentFrequency = 0;  // 当前频率，单位Hz
    private const float maxFrequency = 6000.0f;  // 最大频率，单位Hz
    private const float maxCurrent = 1500.0f;  // 最大电流，单位A
    private float targetFrequency = 0;  // 目标频率，单位Hz

    private ushort final_current = 0;
    
    private string response_log = "";

    private bool comm_state = false;

    // 静态构造函数，初始化串口
    static ModbusRtuSlave()
    {
        // 串口未打开时初始化
        serialPort = new SerialPort();
        serialPort.DataReceived += SerialPort_DataReceived; // 注册 DataReceived 事件
    }

    // 设置串口参数的接口，外部调用此方法设置串口
    public static void SetSerialPortSettings(string portName = "COM3", int baudRate = 9600, Parity parity = Parity.None, int dataBits = 8, StopBits stopBits = StopBits.One)
    {
        if (!isPortOpen)
        {
            serialPort.PortName = portName;
            serialPort.BaudRate = baudRate;
            serialPort.Parity = parity;
            serialPort.DataBits = dataBits;
            serialPort.StopBits = stopBits;
            serialPort.Open();
            isPortOpen = true;  // 标记串口已打开
        }
        else
        {
            Console.WriteLine("串口已经打开，无法更改设置。");
        }
    }

    // 构造函数，初始化每个从站
    public ModbusRtuSlave(byte slaveAddress)
    {
        this.slaveAddress = slaveAddress;
        this.holdingRegisters = new ushort[0x3010];  // 默认有 100 个寄存器
        this.currentFrequency = 0;  // 初始化频率为 0Hz
        slaveInstances[slaveAddress] = this;
    }

    // 启动 Modbus RTU 从站（只启动一个串口线程）
    public static void Start()
    {
        // 启动 Modbus 处理线程（注意，这里没有使用 Thread.Sleep）
        Console.WriteLine("Modbus RTU Slave started...");
    }

    // 设置保持寄存器值（包括频率的目标值）
    public void SetHoldingRegister(ushort address, ushort value)
    {
        Console.WriteLine("Setting Addr:" + address.ToString());
        if (address == 0x2001)  // 设置频率值
        {
            Console.WriteLine("Setting Target Frequency:" + value.ToString());
            this.targetFrequency = value;  // 设置目标频率
            Task.Run(() => GradualFrequencyChange());  // 启动频率逐步变化的任务
        }
        else
        {
            // 设置其他寄存器值的逻辑
            holdingRegisters[address] = value;
        }
    }

    public ushort GetHoldingRegister(ushort address)
    {
        return holdingRegisters[address];
    }

    public string GetCurentLogString()
    {
        return response_log;
    }

    public bool GetCommState()
    {
        return comm_state;
    }

    // 逐步变化频率（模拟真实场景）
    private void GradualFrequencyChange()
    {
        float startFrequency = currentFrequency;
        float changeDuration = 1.2f;

        if (Math.Abs(targetFrequency - startFrequency) > 20 * 100)
        {
            changeDuration = 6.5f;
        }
        else if (Math.Abs(targetFrequency - startFrequency) > 10 * 100 && Math.Abs(targetFrequency - startFrequency) < 20 * 100)
        {
            changeDuration = 4.5f;
        }
        else if (Math.Abs(targetFrequency - startFrequency) > 5 * 100 && Math.Abs(targetFrequency - startFrequency) < 10 * 100)
        {
            changeDuration = 3.3f;
        }
        else if (Math.Abs(targetFrequency - startFrequency) > 250 && Math.Abs(targetFrequency - startFrequency) < 500)
        {
            changeDuration = 2.5f;
        }
        else
        {
            changeDuration = 1.3f;
        }

        float frequencyChangeRate = (targetFrequency - startFrequency) / changeDuration; // 每秒变化频率

        // 逐步变化频率
        for (float t = 0; t < changeDuration; t += 0.1f) // 每 0.1 秒变化一次
        {
            currentFrequency = startFrequency + frequencyChangeRate * t;
            Console.WriteLine($"Current Frequency: {currentFrequency:F2} Hz");

            // 在 Modbus 保持寄存器地址 0x3000 返回频率值（模拟返回）
            holdingRegisters[0x3000] = (ushort)currentFrequency;
            holdingRegisters[0x3004] = (ushort)this.GetCurrent();

            // 模拟每 0.1 秒的时间间隔
            Task.Delay(88).Wait();
        }

        // 确保最终频率与目标频率一致
        currentFrequency = targetFrequency;

        holdingRegisters[0x3000] = (ushort)currentFrequency;
        holdingRegisters[0x3004] = (ushort)this.GetCurrent();

        final_current = (ushort)this.GetCurrent();

        Console.WriteLine($"Final Frequency: {currentFrequency:F2} Hz");
        Console.WriteLine($"Final Current: {currentFrequency:F2} A");
    }

    // 获取当前频率
    public float GetFrequency()
    {
        return currentFrequency;
    }

    // 获取当前电流
    public float GetCurrent()
    {
        Random random = new Random();
        // 生成一个在 -50 到 50 之间的随机数
        int randomNumber = random.Next(-3, 4);
        // 计算电流，最大频率 60Hz 对应最大电流 15A
        return (currentFrequency / maxFrequency) * maxCurrent + (float)(randomNumber);
    }

    public ushort GetFinalCurrent()
    {
        return final_current;
    }

    // DataReceived 事件处理方法
    private static void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        // 每次有数据到达时触发
        int bytesToRead = serialPort.BytesToRead;
        byte[] request = new byte[bytesToRead];
        serialPort.Read(request, 0, bytesToRead);

        byte slaveAddress = request[0];  // 获取请求中的从站地址

        if (slaveInstances.ContainsKey(slaveAddress))  // 如果字典中有该从站实例
        {
            var slave = slaveInstances[slaveAddress];  // 获取对应的从站实例
            byte functionCode = request[1];  // 功能码
            byte[] response = null;

            // 处理 Modbus 读保持寄存器请求（功能码 0x03）
            if (functionCode == 0x03)
            {
                ushort startingAddress = (ushort)((request[2] << 8) + request[3]);
                ushort quantity = (ushort)((request[4] << 8) + request[5]);

                // 处理该从站的保持寄存器读取请求
                response = slave.HandleReadHoldingRegisters(startingAddress, quantity);
            }
            // 处理 Modbus 写单个寄存器请求（功能码 0x06）
            else if (functionCode == 0x06)
            {
                ushort registerAddress = (ushort)((request[2] << 8) + request[3]);
                ushort registerValue = (ushort)((request[4] << 8) + request[5]);

                // 处理写单个寄存器
                response = slave.HandleWriteSingleRegister(registerAddress, registerValue);
            }

            if (response != null)
            {
                slave.comm_state = !slave.comm_state;

                serialPort.Write(response, 0, response.Length);
                Console.WriteLine("Sent response: " + BitConverter.ToString(response));
                slave.response_log = "Sent response: " + BitConverter.ToString(response);
            }
        }
    }

    // 处理读保持寄存器（功能码 0x03）
    private byte[] HandleReadHoldingRegisters(ushort startingAddress, ushort quantity)
    {
        if (startingAddress >= holdingRegisters.Length || startingAddress + quantity > holdingRegisters.Length)
        {
            Console.WriteLine("Invalid register range.");
            return null;  // 无效的寄存器地址范围
        }

        byte[] response = new byte[5 + 2 * quantity]; // 响应长度：功能码 + 字节数 + 寄存器数据

        response[0] = slaveAddress;  // 从站地址
        response[1] = 0x03;  // 功能码（0x03：读取保持寄存器）  
        response[2] = (byte)(2 * quantity); // 数据字节数（每个寄存器 2 字节）

        // 写入寄存器值
        for (int i = 0; i < quantity; i++)
        {
            ushort registerValue = holdingRegisters[startingAddress + (ushort)i];
            response[3 + 2 * i] = (byte)(registerValue >> 8);  // 高字节
            response[4 + 2 * i] = (byte)(registerValue & 0xFF);  // 低字节
        }

        // 校验（CRC 检查）
        byte[] crc = CalculateCRC(response.Take(response.Length - 2).ToArray()); // CRC 检查
        response[response.Length - 2] = crc[0];
        response[response.Length - 1] = crc[1];

        // 比较计算出的 CRC 和接收到的 CRC 是否一致
        if (!ValidateCRC(response))
        {
            Console.WriteLine("CRC mismatch!");
            this.response_log = "Read Holding Registers CRC mismatch!";
            return null;  // CRC 不匹配时返回 null
        }

        return response;
    }

    // 处理写单个保持寄存器（功能码 0x06）
    private byte[] HandleWriteSingleRegister(ushort registerAddress, ushort registerValue)
    {
        if (registerAddress >= holdingRegisters.Length)
        {
            Console.WriteLine("Invalid register address.");
            return null;  // 无效的寄存器地址
        }

        // 更新寄存器值
        holdingRegisters[registerAddress] = registerValue;

        byte[] response = new byte[6]; // 响应长度：从站地址 + 功能码 + 寄存器地址 + 寄存器值 + CRC

        response[0] = slaveAddress;  // 从站地址
        response[1] = 0x06;  // 功能码（0x06：写单个寄存器）
        response[2] = (byte)(registerAddress >> 8);  // 寄存器地址高字节
        response[3] = (byte)(registerAddress & 0xFF);  // 寄存器地址低字节
        response[4] = (byte)(registerValue >> 8);  // 寄存器值高字节
        response[5] = (byte)(registerValue & 0xFF);  // 寄存器值低字节

        // 校验（CRC 检查）
        byte[] crc = CalculateCRC(response.Take(response.Length - 2).ToArray()); // CRC 检查
        response[response.Length - 2] = crc[0];
        response[response.Length - 1] = crc[1];


        // 比较计算出的 CRC 和接收到的 CRC 是否一致
        if (!ValidateCRC(response))
        {
            Console.WriteLine("CRC mismatch!");
            this.response_log = "Write Single Register CRC mismatch!";
            return null;  // CRC 不匹配时返回 null
        }

        SetHoldingRegister(registerAddress, registerValue);

        return response;
    }

    // CRC 校验方法（假设使用的是 Modbus CRC16）
    private byte[] CalculateCRC(byte[] data)
    {
        ushort crc = 0xFFFF;

        foreach (byte byteData in data)
        {
            crc ^= byteData;

            for (int i = 0; i < 8; i++)
            {
                if ((crc & 0x0001) != 0)
                {
                    crc >>= 1;
                    crc ^= 0xA001;
                }
                else
                {
                    crc >>= 1;
                }
            }
        }

        return new byte[] { (byte)(crc & 0xFF), (byte)((crc >> 8) & 0xFF) };
    }

    // 校验计算出来的 CRC 是否和响应中的 CRC 一致
    private bool ValidateCRC(byte[] response)
    {
        // 提取响应中存储的 CRC 值
        byte[] receivedCRC = new byte[] { response[response.Length - 2], response[response.Length - 1] };

        // 计算实际的 CRC 值
        byte[] calculatedCRC = CalculateCRC(response.Take(response.Length - 2).ToArray());

        // 比较计算出的 CRC 和响应中的 CRC
        return receivedCRC.SequenceEqual(calculatedCRC);
    }
}


