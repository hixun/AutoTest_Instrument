
// PI_SerialDlg.cpp : implementation file
//

#include "stdafx.h"
#include "PI_Serial.h"
#include "PI_SerialDlg.h"
#include "afxdialogex.h"
#include <string>
#include <stdlib.h>
#include <atlconv.h>

#pragma comment (lib, "visa32.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CPI_SerialDlg dialog

CPI_SerialDlg::CPI_SerialDlg(CWnd* pParent /*=NULL*/)    //构造函数
	: CDialogEx(CPI_SerialDlg::IDD, pParent)
	, m_dRealResult(0)
	, m_dRealFreq(0)
	, m_dRealRange(0)
	, m_dStartFreq(0)
	, m_dStopFreq(0)
	, m_dStartRange(0)
	, m_dStopRange(0)
	, real_range(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	//对数据初始化
	m_dStartFreq = 10;
	m_dStopFreq = 3000;
	m_dStartRange = -50;
	m_dStopRange = +10;

	m_dRealFreq = 50;
	m_dRealRange = 10;

}

void CPI_SerialDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_SHOW_RESULT, m_dRealResult);
	DDX_Text(pDX, IDC_SET_FREQ, m_dRealFreq);
	DDX_Text(pDX, IDC_SET_RANGE, m_dRealRange);
	DDX_Control(pDX, IDC_COMBO_COM, m_Combo_COM);
	DDX_Control(pDX, IDC_COMBO_BAUD, m_Combo_Baud);
	DDX_Control(pDX, IDC_MSCOMM, m_mscomm);
	DDX_Text(pDX, IDC_START_FREQ, m_dStartFreq);
	DDX_Text(pDX, IDC_STOP_FREQ, m_dStopFreq);
	DDX_Text(pDX, IDC_START_RANGE, m_dStartRange);
	DDX_Text(pDX, IDC_STOP_RANGE, m_dStopRange);
	DDX_Control(pDX, IDC_STEP_FREQ, m_Combo_Stepfreq);
	DDX_Control(pDX, IDC_STEP_RANGE, m_Combo_Steprange);
	DDX_Control(pDX, IDC_PROGRESS1, m_progress);
}

BEGIN_MESSAGE_MAP(CPI_SerialDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_OPEN_PORT, &CPI_SerialDlg::OnBnClickedButtonOpenPort)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_OPEN_INSTRUMENT, &CPI_SerialDlg::OnBnClickedOpenInstrument)
	ON_BN_CLICKED(IDC_INT_ZERO, &CPI_SerialDlg::OnBnClickedIntZero)
	ON_BN_CLICKED(IDC_BUTTON_SINGLE_TEST, &CPI_SerialDlg::OnBnClickedButtonSingleTest)
	ON_BN_CLICKED(IDC_START_TEST, &CPI_SerialDlg::OnBnClickedStartTest)
END_MESSAGE_MAP()


// CPI_SerialDlg message handlers

/************初始化处理函数************/
BOOL CPI_SerialDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 初始化函数，可以添加界面打开时默认显示的内容

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	m_progress.SetRange(1, 80);		//进度条设置初始化
	m_progress.SetStep(1);
	m_progress.SetPos(0);

	try{

		CoInitialize(NULL);
		m_pConnection = _ConnectionPtr(__uuidof(Connection));
		m_pConnection->ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=test1.accdb;";
		m_pConnection->Open("", "", "", adConnectUnspecified);
	}
	catch (_com_error e){

		AfxMessageBox(_T("数据库连接失败！"));
	}

	// 串口选择组合框
	CString str;
	int i;
	for (i = 0; i<15; i++)
	{
		str.Format(_T("COM %d"), i + 1);
		m_Combo_COM.InsertString(i, str);
	}
	m_Combo_COM.SetCurSel(5);//预置COM6口

	//波特率选择组合框
	CString str1[] = { _T("300"), _T("600"), _T("1200"), _T("2400"), _T("4800"), _T("9600"),
		_T("19200"), _T("38400"), _T("43000"), _T("56000"), _T("57600"), _T("115200") };
	for (int i = 0; i<12; i++)
	{
		int judge_tf = m_Combo_Baud.AddString(str1[i]);
		if ((judge_tf == CB_ERR) || (judge_tf == CB_ERRSPACE))
			MessageBox(_T("build baud error!"));
	}
	m_Combo_Baud.SetCurSel(11);//预置波特率为"115200"

	//频率步进选择组合框
	CString str2[] = { _T("0.001"), _T("0.01"), _T("0.1"), _T("1"), _T("10"), _T("20"), _T("40"), _T("50"), _T("100"), _T("200") };
	for (int i = 0; i<10; i++)
	{
		int judge_tf = m_Combo_Stepfreq.AddString(str2[i]);
		if ((judge_tf == CB_ERR) || (judge_tf == CB_ERRSPACE))
			MessageBox(_T("build stepfreq error!"));
	}
	m_Combo_Stepfreq.SetCurSel(6);//预置频率步进为"40"

	//幅度步进选择组合框
	CString str3[] = { _T("0.05"), _T("0.1"), _T("0.5"), _T("1"), _T("2"), _T("5"), _T("10") };
	for (int i = 0; i<6; i++)
	{
		int judge_tf = m_Combo_Steprange.AddString(str3[i]);
		if ((judge_tf == CB_ERR) || (judge_tf == CB_ERRSPACE))
			MessageBox(_T("build steprange error!"));
	}
	m_Combo_Steprange.SetCurSel(4);//预置幅度步进为"2"

	num_freq = 0;
	num_range = 0;

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CPI_SerialDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CPI_SerialDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CPI_SerialDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


/************点击 打开串口 按钮时************/
void CPI_SerialDlg::OnBnClickedButtonOpenPort()
{
	// TODO: Add your control notification handler code here
	CString str, str1, n;
	GetDlgItemText(IDC_BUTTON_OPEN_PORT, str);
	CWnd *hl;
	hl = GetDlgItem(IDC_BUTTON_OPEN_PORT);	//指向打开串口按钮控件

	if (!m_mscomm.get_PortOpen())
	{
		m_Combo_Baud.GetLBText(m_Combo_Baud.GetCurSel(), str1);		//将波特率的值放到str1里
		str1 = str1 + ',' + 'n' + ',' + '8' + ',' + '1';		//8位数据，一个校验，将这些放到一个字符串中
		m_mscomm.put_CommPort((m_Combo_COM.GetCurSel() + 1));	//选择串口
		m_mscomm.put_InputMode(1);			//设置输入方式为二进制方式
		m_mscomm.put_Settings(str1);		//波特率为（波特率组Á合框）无校验，8数据位，1个停止位
		m_mscomm.put_InputLen(2048);		//设置当前接收区数据长度为1024
		m_mscomm.put_OutBufferSize(1024);//发送缓冲区 
		m_mscomm.put_RThreshold(1);			//缓冲区一个字符引发事件
		m_mscomm.put_RTSEnable(1);			//设置RT允许

		m_mscomm.put_PortOpen(true);		//打开串口
		if (m_mscomm.get_PortOpen())
		{
			str = _T("关闭串口");
			UpdateData(true);
			hl->SetWindowText(str);			//改变按钮名称为‘’关闭串口”
		}
		CByteArray senddata;
		const int data_zero = 0x0a;
		senddata.Add(data_zero);

		m_mscomm.put_Output(COleVariant(_T("REF:OSC:SOURCE EXT")));	//串口向模块发送数据进行初始化
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("ALC:STATe ON")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("ALC:TEMP:STATe ON")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("PULM:STATe OFF")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("PULM:SWITch ON")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("FREQuency 3.2GHz")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("CALibration:ALCValue:CHANel DERect")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("POWer:ATT 0dB")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("POWer:LEV 10dBm")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("OUTPut ON")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("*IDN?")));
		m_mscomm.put_Output(COleVariant(senddata));
		m_mscomm.put_Output(COleVariant(_T("POWer:LEVel?")));
		m_mscomm.put_Output(COleVariant(senddata));
	}
	else
	{
		m_mscomm.put_PortOpen(false);
		if (str != _T("打开串口"))
		{
			str = _T("打开串口");
			UpdateData(true);
			hl->SetWindowText(str);			//改变按钮名称为打开串口
			KillTimer(1);	// 关闭定时器   
		}
	}
}
BEGIN_EVENTSINK_MAP(CPI_SerialDlg, CDialogEx)
	ON_EVENT(CPI_SerialDlg, IDC_MSCOMM, 1, CPI_SerialDlg::OnCommMscomm, VTS_NONE)
END_EVENTSINK_MAP()


/************串口处理函数，主要用于接收串口数据，暂时没有用到************/
void CPI_SerialDlg::OnCommMscomm()
{
	// TODO: Add your message handler code here
}


/************点击 打开仪器 按钮时************/
void CPI_SerialDlg::OnBnClickedOpenInstrument()
{
	// TODO: Add your control notification handler code here
	CString str;
	GetDlgItemText(IDC_OPEN_INSTRUMENT, str);
	CWnd *hl;
	hl = GetDlgItem(IDC_OPEN_INSTRUMENT);	//指向打开仪器按钮控件

	if ((VI_SUCCESS != viGetAttribute(Power_Sensors_U2000, VI_ATTR_RSRC_CLASS, intfName)))
	{
		//打开总的资源管理器，初始化资源管理器  
		viOpenDefaultRM(&rm);
		//打开指定的USB接口控制的函数发生器  
		viOpen(rm, "USB0::2391::11032::MY48100910::0::INSTR", VI_NULL, VI_NULL, &Power_Sensors_U2000);

		//确认默认的函数发生器命令否以\n结束，这里定义的SCPI语言是必须以\n结尾的  
		ViBoolean termDefaultPower_Sensors_U2000 = VI_FALSE;
		if ((VI_SUCCESS == viGetAttribute(Power_Sensors_U2000, VI_ATTR_RSRC_CLASS, intfName)))
		{
			termDefaultPower_Sensors_U2000 = VI_TRUE;
			str = _T("关闭仪器");
			UpdateData(true);
			hl->SetWindowText(str);			//改变按钮名称为关闭仪器
			KillTimer(1);	// 关闭定时器   
		}
	}
	else
	{
		//关闭到指定的USB接口控制的功率计的连接  
		viClose(Power_Sensors_U2000);
		//关闭总的资源管理器  
		viClose(rm);
		if (str != _T("打开仪器"))
		{
			str = _T("打开仪器");
			UpdateData(true);
			hl->SetWindowText(str);			//改变按钮名称为打开仪器
		}
	}
	UpdateData(false);	
}

/************点击 INT ZERO 按钮时************/
void CPI_SerialDlg::OnBnClickedIntZero()
{
	// TODO: Add your control notification handler code here
	INT_PTR nRes;
	char zero_buf[2] = { 0 };


	//归零操作
	viPrintf(Power_Sensors_U2000, "CAL:ZERO:AUTO\n");
	viPrintf(Power_Sensors_U2000, "*OPC?\n");

	viScanf(Power_Sensors_U2000, "%s", zero_buf);

	if (zero_buf[0] == '1')
	{
		// 显示消息对话框   
		nRes = MessageBox(_T("Completed!"), _T("ZERO"), MB_ICONQUESTION);
	}
	else
	{
		// 显示消息对话框   
		nRes = MessageBox(_T("Not completed!"), _T("ZERO"), MB_ICONQUESTION);
	}
}

/************点击 单次测试 按钮时************/
void CPI_SerialDlg::OnBnClickedButtonSingleTest()
{
	// TODO: Add your control notification handler code here
	CByteArray senddata;
	const int data_zero = 0x0a;
	senddata.Add(data_zero);

	CString COM_Send_freq;	//通过串口要发送的频率
	CString COM_Send_range;	//通过串口要发送的幅度
	CString Single_Power_freq;//设置功率计的频率
	
	UpdateData(TRUE);

	m_dSingle_freq = m_dRealFreq;
	m_dSingle_range = m_dRealRange;

	COM_Send_freq.Format(_T("FREQ %.3lfMHz"), m_dSingle_freq);
	COM_Send_range.Format(_T("POWer:LEVel %.3lfdBm"), m_dSingle_range);


	m_mscomm.put_Output(COleVariant(_T("CALibration:ALCValue:CHANel 10MTO1G")));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(_T("POWer:ATT 0dB")));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(COM_Send_freq));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(COM_Send_range));
	m_mscomm.put_Output(COleVariant(senddata));

	Sleep(1000);

	viPrintf(Power_Sensors_U2000, "INIT:CONT ON\n");
	viPrintf(Power_Sensors_U2000, "FREQ 1000MHz\n");
	viPrintf(Power_Sensors_U2000, "FETC?\n");
	viFlush(Power_Sensors_U2000, VI_READ_BUF_DISCARD);		//放弃缓读区的内容

	viScanf(Power_Sensors_U2000, "%s", buf);

	real_range = strtod(buf, NULL) + 0;
	m_dRealResult = real_range;
//	UpdateData(FALSE);

	//_variant_t RecordsAffected;
	//CString AddSql;
	//AddSql.Format(_T("INSERT INTO userInfo(sendFreq,sendRange,receiveRange) VALUES(%.3lf,%.3lf,%.3lf)"), m_dSingle_freq, m_dSingle_range, m_dRealResult);

	//try{
	//	m_pConnection->Execute((_bstr_t)AddSql, &RecordsAffected, adCmdText);
	//	AfxMessageBox(_T("添加用户成功！"));
	//}
	//catch (_com_error* e){
	//	AfxMessageBox(_T("添加用户失败！"));
	//}

	UpdateData(FALSE);

}

/************点击 开始测试 按钮时************/
void CPI_SerialDlg::OnBnClickedStartTest()
{
	// TODO: Add your control notification handler code here
	CString str;
	GetDlgItemText(IDC_START_TEST, str);
	CWnd *hl;
	hl = GetDlgItem(IDC_START_TEST);	//指向开始测试按钮控件

	CString str_StepFreq;
	CString str_StepRange;



	UpdateData(TRUE);

	m_Combo_Stepfreq.GetLBText(m_Combo_Stepfreq.GetCurSel(), str_StepFreq);		//将步进频率的值放到str_StepFreq里
	m_Combo_Steprange.GetLBText(m_Combo_Steprange.GetCurSel(), str_StepRange);		//将步进幅度的值放到str_StepRange里

	m_dStepFreq = _tstof(str_StepFreq);		//将步进频率转换为double型
	m_dStepRange = _tstof(str_StepRange);	//将步进幅度转换为double型

	num_freq = (int)((m_dStopFreq - m_dStartFreq) / m_dStepFreq);		//计算出频率改变次数
	num_range = (int)((m_dStopRange - m_dStartRange) / m_dStepRange);	//计算出幅度改变次数

	UpdateData(FALSE);
	if (str == _T("开始测试"))
	{
		str = _T("停止测试");
		UpdateData(true);
		hl->SetWindowText(str);			//改变按钮名称为停止测试
		SetTimer(1, 1100, NULL);         // 启动定时器，ID为1，定时时间为200ms 
	}else if (str != _T("开始测试"))
		{
			str = _T("开始测试");
			UpdateData(true);
			hl->SetWindowText(str);			//改变按钮名称为开始测试
			KillTimer(1);	// 关闭定时器   
		}
}

/************定时器时间到达时处理函数************/
void CPI_SerialDlg::OnTimer(UINT_PTR nIDEvent)
{
	double cur_freq;	//当前频率
	double cur_range;

	CByteArray senddata;
	const int data_zero = 0x0a;
	senddata.Add(data_zero);

	CString COM_Send_freq;	//通过串口要发送的频率
	CString COM_Send_range;	//通过串口要发送的幅度

	cur_range = m_dStartRange + count_range *  m_dStepRange;	//当前应该向模块发送的幅度值
	cur_freq = m_dStartFreq + count_freq * m_dStepFreq;			//当前应该向模块发送的频率值

	COM_Send_freq.Format(_T("FREQ %.3lfMHz"), cur_freq);		//将double型幅度数据转换为安泰信模块固定格式
	COM_Send_range.Format(_T("POWer:LEVel %.3lfdBm"), cur_range);

//	m_mscomm.put_Output(COleVariant(COM_Send_freq));	//串口向模块发送数据
//	m_mscomm.put_Output(COleVariant(COM_Send_range));

	m_mscomm.put_Output(COleVariant(_T("CALibration:ALCValue:CHANel 10MTO1G")));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(_T("POWer:ATT 0dB")));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(COM_Send_freq));
	m_mscomm.put_Output(COleVariant(senddata));
	m_mscomm.put_Output(COleVariant(COM_Send_range));
	m_mscomm.put_Output(COleVariant(senddata));

	count_range++;
	if (count_range == num_range + 1)
	{
		count_range = 0;
		count_freq++;
	}
	if (count_freq == num_freq + 1)
	{
		KillTimer(1);	// 关闭定时器   
	}

	Sleep(1000);

	viPrintf(Power_Sensors_U2000, "FREQ %fMHz\n", cur_freq);	//USB功率计指令,将double型频率数据转换为USB功率计固定格式
	viPrintf(Power_Sensors_U2000, "FETC?\n");

	viScanf(Power_Sensors_U2000, "%s", buf);	//读取USB功率计的读数
	viFlush(Power_Sensors_U2000, VI_READ_BUF_DISCARD);		//放弃缓读区的内容

	UpdateData(TRUE);
	m_dRealFreq = cur_freq;		//将发送给模块的频率，幅度，从功率计读取的数据显示在当前状态栏
	m_dRealRange = cur_range;

	real_range = strtod(buf, NULL);
	m_dRealResult = real_range;

	progress_step++;
	if (progress_step == 2)
	{
		m_progress.StepIt();	//更新进度条
		progress_step = 0;
	}

	UpdateData(FALSE);


	_variant_t RecordsAffected;
	CString AddSql;
	AddSql.Format(_T("INSERT INTO data1(sendFreq,sendRange,receiveRange) VALUES(%.3lf,%.3lf,%.3lf)"), m_dRealFreq, m_dRealRange, m_dRealResult);

	try{
		m_pConnection->Execute((_bstr_t)AddSql, &RecordsAffected, adCmdText);
	}
	catch (_com_error* e){
		//		AfxMessageBox(_T("添加用户失败！"));
	}

	CDialogEx::OnTimer(nIDEvent);
}


