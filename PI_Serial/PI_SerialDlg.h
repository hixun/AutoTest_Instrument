
// PI_SerialDlg.h : header file
//

#pragma once
#include "afxwin.h"
#include "mscomm.h"
#include "visa.h"
#include "afxcmn.h"


// CPI_SerialDlg dialog
class CPI_SerialDlg : public CDialogEx
{
// Construction
public:
	CPI_SerialDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_PI_SERIAL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

	//关于仪器的变量
	ViSession rm;
	ViSession Power_Sensors_U2000;
	ViUInt16 io_prot;
	ViUInt16 intfType;
	ViString intfName[512];
	char buf[256];
	double receive;
	int zero_flag ;
	CString m_strRec;
	int m_iShowtest;

	double m_dSingle_freq;	//单次输入的频率
	double m_dSingle_range;	//单次输入的幅度

	double m_dStepFreq;
	double m_dStepRange;

	int num_freq;		//频率共要改变多少次
	int num_range;		//幅度共要改变多少次

	int count_freq;		//频率步进次数
	int count_range;	//幅度步进次数

	int progress_step;	//多少次进度条改变一下
	long change_step;

	double real_range;

	double Offset_Range;	//要补偿的幅度
	double Offset_Range1;	//要补偿的幅度
	double Offset_Range2;	//要补偿的幅度

	_ConnectionPtr m_pConnection;//连接access数据库的链接对象  
	_RecordsetPtr m_pRecordset;//结果集对象  

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	double m_dRealResult;
	double m_dRealFreq;
	double m_dRealRange;
	CComboBox m_Combo_COM;
	CComboBox m_Combo_Baud;
	CMscomm m_mscomm;
	double m_dStartFreq;
	double m_dStopFreq;
	double m_dStartRange;
	double m_dStopRange;
	CComboBox m_Combo_Stepfreq;
	CComboBox m_Combo_Steprange;
	afx_msg void OnBnClickedButtonOpenPort();
	DECLARE_EVENTSINK_MAP()
	void OnCommMscomm();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnBnClickedOpenInstrument();
	afx_msg void OnBnClickedIntZero();
	afx_msg void OnBnClickedButtonSingleTest();
	afx_msg void OnBnClickedStartTest();
	CProgressCtrl m_progress;
	double m_pOffsetRange;
};
