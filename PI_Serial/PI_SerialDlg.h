
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

	//���������ı���
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

	double m_dSingle_freq;	//���������Ƶ��
	double m_dSingle_range;	//��������ķ���

	double m_dStepFreq;
	double m_dStepRange;

	int num_freq;		//Ƶ�ʹ�Ҫ�ı���ٴ�
	int num_range;		//���ȹ�Ҫ�ı���ٴ�

	int count_freq;		//Ƶ�ʲ�������
	int count_range;	//���Ȳ�������

	int progress_step;	//���ٴν������ı�һ��
	long change_step;

	double real_range;

	double Offset_Range;	//Ҫ�����ķ���
	double Offset_Range1;	//Ҫ�����ķ���
	double Offset_Range2;	//Ҫ�����ķ���

	_ConnectionPtr m_pConnection;//����access���ݿ�����Ӷ���  
	_RecordsetPtr m_pRecordset;//���������  

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
