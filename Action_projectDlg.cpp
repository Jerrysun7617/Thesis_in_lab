
// Action_projectDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Action_project.h"
#include "Action_projectDlg.h"
#include "afxdialogex.h"

#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include <iostream>
#include <string.h>
#include <math.h>

#define PI 3.14159265358979323846
#define ACCEL
#define GYRO
//#define GESTURE_THREE

#ifdef _DEBUG
#define new DEBUG_NEW
#endif




//define the variables here
CApplication ExcelApp;
//accelerometer
CWorkbooks wbsMyBooks_accel;
CWorkbook wbMyBook_accel;
CWorksheets wsMySheets_accel;
CWorksheet wsMySheet_accel;
CRange rgMyRge_accel;
//gyroscope
CWorkbooks wbsMyBooks_gyro;
CWorkbook wbMyBook_gyro;
CWorksheets wsMySheets_gyro;
CWorksheet wsMySheet_gyro;
CRange rgMyRge_gyro;


//Transform Domain Algorithm
struct Domain_Num {
	int numi;
	int numi_last;
};
struct Domain_Num Domain = { 0,0 };

//********* action Algorithms, 
struct Action_angle_strict {

	public:
		// for the gyro angle
		double gyro_x[2];
		//int gyro_x_max;
		double gyro_y[2];
		//int gyro_y_max;
		double gyro_z[2];
		//int gyro_z_max;

		//for the accel angle
		double accel_x[2];
		double accel_y[2];
		double accel_z[2];
};
/*
struct Action_angle_strict action_123[2] = {
	{ -30,30,-10,10,-10,10,-15,15,-10,13,-63,63 },  // action1  control Z-axis
	{ -10, 10, -10, 10, -30, 30,  -63, 63,  -10,30,  -10, 30 }, // action2 X-aixs
};
*/
struct Action_angle_strict action_123[3] = {
		{ -150,150,-60,60,-60,60,    -15,15,-13,15,-58,58 },  // action1  control Z-axis
		{-80, 60, -60, 60, -150, 150,     -58, 58,  -20,30,  -40, 25}, // action2 X-aixs
#ifdef GESTURE_THREE
		{ -15, 15, -15, 15, -20, 20, -40, 40, 0, 10, -40, 40 }, // action3 Y-axis
		//{ -15, 15, -15, 15, -20,   20, 40, 20, 70, 50, 70 } // action3 Y-axis
#endif
	};

//********* for recording the action committed times.
class Action_judge {
private:
	bool flag_action1;
	bool flag_action2;
	bool flag_action3;

	int act_1_times;
	int act_2_times;
	int act_3_times;

public:

	int Total_a1;
	int Total_a2;
	int Total_a3;

	 Action_judge(void) {
		flag_action1 = false;
		flag_action2 = false;
		flag_action3 = false;

		act_1_times = 0;
		act_2_times = 0;
		act_3_times = 0;

		Total_a1 = 0;
		Total_a2 = 0;
		Total_a3 = 0;
	}
	void initialisation(bool f1, bool f2, bool f3, int a1, int a2, int a3) {
		flag_action1 = f1;
		flag_action2 = f2;
		flag_action3 = f3;

		act_1_times = a1;
		act_2_times = a2;
		act_3_times = a3;
	}
};

struct Accel_data {
	double Daccel_x;
	double Daccel_y;
	double Daccel_z;

	double A_angle_x;
	double A_angle_y;
	double A_angle_z;
};
struct Gyro_data {
	double Dgyro_x;
	double Dgyro_y;
	double Dgyro_z;


	double G_angle_x;
	double G_angle_y;
	double G_angle_z;
};
struct Gyro_data Gyro;
struct Actions {
	int Act_01;
	int Act_02;
	int Act_03;
};
class Jerry_Class {

	public:
		void Action_recognition(struct Accel_data Accel, struct Gyro_data Gyro) {


		}
};

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CAction_projectDlg �Ի���



CAction_projectDlg::CAction_projectDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_ACTION_PROJECT_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CAction_projectDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAction_projectDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_EN_CHANGE(IDC_EDIT3, &CAction_projectDlg::OnEnChangeEdit3)
	ON_EN_CHANGE(IDC_EDIT1, &CAction_projectDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT2, &CAction_projectDlg::OnEnChangeEdit2)
	ON_EN_CHANGE(IDC_EDIT5, &CAction_projectDlg::OnEnChangeEdit5)
	ON_EN_CHANGE(IDC_EDIT4, &CAction_projectDlg::OnEnChangeEdit4)
	ON_BN_CLICKED(IDC_BUTTON1, &CAction_projectDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CAction_projectDlg ��Ϣ�������

BOOL CAction_projectDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CAction_projectDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CAction_projectDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CAction_projectDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CAction_projectDlg::OnEnChangeEdit3()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CAction_projectDlg::OnEnChangeEdit1()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CAction_projectDlg::OnEnChangeEdit2()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CAction_projectDlg::OnEnChangeEdit5()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CAction_projectDlg::OnEnChangeEdit4()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}
void Get_gyro_angle(struct Gyro_data &Gyro, struct Accel_data Accel) {
	//double atan(double x)   //����arctan(x)��ֵ
	double K1 = 0.98;
	double K2 = 0.02;

	Gyro.G_angle_x = K1 *(Gyro.G_angle_x+ Gyro.Dgyro_x) + K2*Accel.Daccel_x;

	Gyro.G_angle_y = K1 *(Gyro.G_angle_y + Gyro.Dgyro_y) + K2*Accel.Daccel_y;

	Gyro.G_angle_z = K1 *(Gyro.G_angle_z + Gyro.Dgyro_z) + K2*Accel.Daccel_z;
}

void Get_accel_angle(struct Accel_data &Accel) {
	//double atan(double x)   //����arctan(x)��ֵ
	//doublex=4.0,result;  
	//result = sqrt(x);//result*result=x 
	double Sqrt_x = sqrt(Accel.Daccel_y * Accel.Daccel_y +Accel.Daccel_z * Accel.Daccel_z);
	double Sqrt_y = sqrt(Accel.Daccel_x * Accel.Daccel_x + Accel.Daccel_z * Accel.Daccel_z);
	double Sqrt_z = sqrt(Accel.Daccel_x * Accel.Daccel_x + Accel.Daccel_y * Accel.Daccel_y);

	Accel.A_angle_x = (atan(Accel.Daccel_x / Sqrt_x) * 180) / PI;

	Accel.A_angle_y = (atan(Accel.Daccel_y / Sqrt_y) * 180) / PI;

	Accel.A_angle_z = (atan(Accel.Daccel_z / Sqrt_z) * 180) / PI;

	// set up a coefficient ���ϱߵĺ����Ա�ͼ���Կ�������ϵ��ȡ0.92ʱ���Ƕȷ�Χ��������-45~45�ȡ�
	//https://www.arduino.cn/thread-18392-1-1.html
	double Ks = 0.92;
	//double Ks = 1.2;
	Accel.A_angle_x *= Ks;

	Accel.A_angle_y *= Ks;

	Accel.A_angle_z *= Ks;

}
void CAction_projectDlg::Gyro_file_Init(void) {
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	LPDISPATCH lpDisp_gyro = NULL;

	//create new xls file 
	//����һ��EXCELӦ�ó���ʵ��
#ifndef ACCEL
  	if (!ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("����Excel����ʧ��"));
		exit(1);
	}
#endif
	//G:\\123.xlsx"
	ExcelApp.put_Visible(false);   //ʹ����ɼ�
	ExcelApp.put_UserControl(FALSE);
	
	//������������������Ӧ�ó�����������
	wbsMyBooks_gyro.AttachDispatch(ExcelApp.get_Workbooks());

	CEdit* pMessage_gyro;
	pMessage_gyro = (CEdit*)GetDlgItem(IDC_EDIT2);
	CString strFile_gyro;
	pMessage_gyro->GetWindowText(strFile_gyro);

	if (strFile_gyro.GetLength() == 0)
		return;

	try
	{
		//��һ��������   for accelerometer
		lpDisp_gyro = wbsMyBooks_gyro.Open(strFile_gyro, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
		//������������ʹ򿪵Ĺ���������  
		wbMyBook_gyro.AttachDispatch(lpDisp_gyro);
	}
	catch (...)
	{
		return;
	}

	// ��ù�������������ʵ��
	wsMySheets_gyro.AttachDispatch(wbMyBook_gyro.get_Sheets());

	//��һ�����������������������һ��  
	CString strSheetName_gyro = _T("Sheet1");

	try
	{
		//��һ�����еĹ�����(wooksheet)  
		lpDisp_gyro = wsMySheets_gyro.get_Item(_variant_t(strSheetName_gyro));
		wsMySheet_gyro.AttachDispatch(lpDisp_gyro);
	}
	catch (...)
	{
		lpDisp_gyro = wsMySheets_gyro.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		wsMySheet_gyro.AttachDispatch(lpDisp_gyro);
		wsMySheet_gyro.put_Name(strSheetName_gyro);     //������������  
	}


	//gyro
	//���õ�Ԫ��ķ�ΧΪȫ������  
	lpDisp_gyro = wsMySheet_gyro.get_Cells();
	rgMyRge_gyro.AttachDispatch(lpDisp_gyro);
}
void CAction_projectDlg::Accel_file_Init(void) {
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	LPDISPATCH lpDisp_accel = NULL;

	//create new xls file 
	//����һ��EXCELӦ�ó���ʵ��
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("����Excel����ʧ��"));
		exit(1);
	}
	//G:\\My thesis\\Excel Data\\Action1_20\\A1_20_gyro_C.xlsx"
	//G:\\My thesis\\Excel Data\\Action1_20\\A1_20_accel_C.xlsx"
	ExcelApp.put_Visible(false);   //ʹ����ɼ�
	ExcelApp.put_UserControl(FALSE);

	//������������������Ӧ�ó�����������
	wbsMyBooks_accel.AttachDispatch(ExcelApp.get_Workbooks());

	CEdit* pMessage_accel;
	pMessage_accel = (CEdit*)GetDlgItem(IDC_EDIT1);
	CString strFile_accel;
	pMessage_accel->GetWindowText(strFile_accel);

	if (strFile_accel.GetLength() == 0)
		return;

	try
	{
		//��һ��������   for accelerometer
		lpDisp_accel = wbsMyBooks_accel.Open(strFile_accel, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
		//������������ʹ򿪵Ĺ���������  
		wbMyBook_accel.AttachDispatch(lpDisp_accel);
	}
	catch (...)
	{
		return;
		lpDisp_accel = wbsMyBooks_accel.Add(vtMissing);
		wbMyBook_accel.AttachDispatch(lpDisp_accel);
	}


	// ��ù�������������ʵ��
	wsMySheets_accel.AttachDispatch(wbMyBook_accel.get_Sheets());


	//��һ�����������������������һ��  
	CString strSheetName_accel = _T("Sheet1");

	try
	{
		//��һ�����еĹ�����(wooksheet)  
		lpDisp_accel = wsMySheets_accel.get_Item(_variant_t(strSheetName_accel));
		wsMySheet_accel.AttachDispatch(lpDisp_accel);
		//��һ�����еĹ�����(wooksheet)  
	}
	catch (...)
	{  
	}

	//accel
	//���õ�Ԫ��ķ�ΧΪȫ������  
	lpDisp_accel = wsMySheet_accel.get_Cells();
	rgMyRge_accel.AttachDispatch(lpDisp_accel);
}
int Action123_Judge(struct Accel_data Accel, struct Gyro_data Gyro) {
	bool flag[3] = {false,false,false};

	for (int i = 0; i < 2; i++) {
		
		if (Gyro.G_angle_x > action_123[i].gyro_x[1] || Gyro.G_angle_x <  action_123[i].gyro_x[0])
			continue;
		if (Gyro.G_angle_y > action_123[i].gyro_y[1] || Gyro.G_angle_y < action_123[i].gyro_y[0])
			continue;
		if (Gyro.G_angle_z > action_123[i].gyro_z[1] || Gyro.G_angle_z < action_123[i].gyro_z[0])
			continue;
		
		//G:\\My thesis\\Excel Data Analysis\\GY Analysis\\A1_20\\A1_Accel_20.xlsx
		if (i == 0)
		{
			//judge the Z first for Action 1 and than Y and X
			//{ -30, 30, -10, 10, -10, 10, -15, 15, -15, 15, -60, 60 },  // action1  control Z-axis
			// Z should be <-55 or >55 
			//Y in -20  -- 20 
			//X in -20  -- 20
			if (Accel.A_angle_z <  action_123[i].accel_z[0] || Accel.A_angle_z >action_123[i].accel_z[1])
					if ((Accel.A_angle_y > action_123[i].accel_y[0] && Accel.A_angle_y < action_123[i].accel_y[1]))
						if (Accel.A_angle_x > action_123[i].accel_x[0] && Accel.A_angle_x < action_123[i].accel_x[1])
							flag[i] = true;
		}
		//G:\\My thesis\\Excel Data Analysis\\GY Analysis\\A2_20\\A2_Accel_20.xlsx
		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\GY Analysis\\A2_20\\A2_Accel_20.xlsx
		if (i == 1) {
			// {-10, 10, -10, 10, -30, 30,  -70, 70,  -15,15,  -15, 15}, // action2 X-aixs
			// fix X_axis (-70, 70)  and then Y ��-15�� 15�� Z��-15,15��

			if (Accel.A_angle_x <  action_123[i].accel_x[0] || Accel.A_angle_x >action_123[i].accel_x[1])
				if ((Accel.A_angle_y > action_123[i].accel_y[0] && Accel.A_angle_y < action_123[i].accel_y[1]))
					if (Accel.A_angle_z > action_123[i].accel_z[0] && Accel.A_angle_z < action_123[i].accel_z[1])
						flag[i] = true;
		}
		//{ -15, 15, -15, 15, -20, 20,     -40, 40, -10, 10, 40, 40 }, // action3 Y-axis
		//G:\\My thesis\\Excel Data Analysis\\GY Analysis\\A3_20\\A3_Accel_20.xlsx
#ifdef GESTURE_THREE
		if (i == 2) {
			if (Accel.A_angle_y  > action_123[i].accel_y[0] && Accel.A_angle_y < action_123[i].accel_y[1])
				if ((Accel.A_angle_x < action_123[i].accel_x[0] || Accel.A_angle_x > action_123[i].accel_x[1]))
					if (Accel.A_angle_z < action_123[i].accel_z[0] || Accel.A_angle_z > action_123[i].accel_z[1])
						flag[i] = true;
		}
#endif	
			//if (flag[i] == false)
			/*
				if ((Accel.A_angle_z > 0 - action_123[i].accel_z[1] && Accel.A_angle_z < 0 - action_123[i].accel_z[0]) || (Accel.A_angle_z > action_123[i].accel_z[0] && Accel.A_angle_z < action_123[i].accel_z[1]))
				{
					if ((Accel.A_angle_y > 0 - action_123[i].accel_y[1] && Accel.A_angle_y < 0 - action_123[i].accel_y[0]) || (Accel.A_angle_y > action_123[i].accel_y[0] && Accel.A_angle_y < action_123[i].accel_y[1]))
					{
						if ((Accel.A_angle_x > 0 - action_123[i].accel_x[1] && Accel.A_angle_x < 0 - action_123[i].accel_x[0]) || (Accel.A_angle_x > action_123[i].accel_x[0] && Accel.A_angle_x < action_123[i].accel_x[1]))
							flag[i] = true;
					}
				}
				*/
	
	}


	int num = 0;
	int act = 0;
	for (int i = 0; i < 2; i++) {
		if (flag[i]) {
			num++;
			act = i + 1;
		}
	}
	if (num > 1) //more than one actions identified, it is confused
		act = 0;

	return act;
}

struct action_flag {
	int previous_time;
	int last_times[3];
	bool flag[3];
};
struct action_flag aflag = { 111,{0,0,0},{false,false,false} };
int  Action_Calclulation( struct Accel_data Accel, struct Gyro_data Gyro) {

	int act = 0;
	 act = Action123_Judge(Accel, Gyro);
	if (aflag.previous_time == 111)
		aflag.previous_time = act;

	switch(act){
		case 0:
			act = 0;
			aflag.last_times[0] = 0;
			aflag.last_times[1] = 0;
			aflag.last_times[2] = 0;
			aflag.previous_time = 0;
			break;
		case 1:
			if (aflag.last_times[1] != 0 || aflag.last_times[2] != 0)
			{
				aflag.last_times[0] = 0;
				aflag.last_times[1] = 0;
				aflag.last_times[2] = 0;
				aflag.previous_time = 0;
				act = 0;
			}
			else {
				if (aflag.previous_time == 4)
					aflag.last_times[0]++;
				aflag.previous_time++;


				if (aflag.last_times[0] == 1) {
					
					aflag.last_times[0] = 0;
					// bewtween 0.04s and 250s
					if ((Domain.numi - Domain.numi_last) > 40 && (Domain.numi - Domain.numi_last) < 250)
						act = 1;
					else
						act = 0;
						
					Domain.numi_last = Domain.numi;
				}
				else {
					act = 0;
					//aflag.last_times[0] = 0;
				}
			}
			break;
		case 2:
			if (aflag.last_times[0] != 0 || aflag.last_times[2] != 0)
			{
				aflag.last_times[0] = 0;
				aflag.last_times[1] = 0;
				aflag.last_times[2] = 0;
				aflag.previous_time = 0;
				act = 0;
			}
			else {
				if (aflag.previous_time == 4)
					aflag.last_times[1]++;
				aflag.previous_time++;

				if (aflag.last_times[1] == 1) {
					aflag.last_times[1] = 0;

					// bewtween 0.08s and 250s
					if ((Domain.numi - Domain.numi_last) > 40 && (Domain.numi - Domain.numi_last) < 250)
						act = 2;
					else
						act = 0;

					Domain.numi_last = Domain.numi;
				}
				else {
					act = 0;
				}
			}
			break;
#ifdef GESTURE_THREE
		case 3:
			if (aflag.last_times[0] != 0 || aflag.last_times[1] != 0)
			{
				aflag.last_times[0] = 0;
				aflag.last_times[1] = 0;
				aflag.last_times[2] = 0;
				aflag.previous_time = 0;
				act = 0;
			}else {
				if (aflag.previous_time == 3)
					aflag.last_times[2]++;
				aflag.previous_time++;

				if (aflag.last_times[2] == 1) {
					aflag.last_times[2] = 0;
					act = 3;
				}
				else {
					act = 0;
				}
			}
			break;
#endif
		default:
			act = 0;
			aflag.last_times[0] = 0;
			aflag.last_times[1] = 0;
			aflag.last_times[2] = 0;
			aflag.previous_time = 0;
			break;
	}
	return act;
}
void CAction_projectDlg::OnBnClickedButton1()
{
#ifdef ACCEL
	//initialize accel
	Accel_file_Init();
#endif

	//initialize gyroscope
	Gyro_file_Init();

	VARIANT Accele_x, Accele_y, Accele_z;
	CString str(_T(""));
	char cMsg_x[20];
	CString str_x;
	_bstr_t bMsg_x;
	struct Accel_data Accel = { 0,0,0,0,0,0 };

	VARIANT Gyro_x, Gyro_y, Gyro_z;
	Gyro  = {0.0};


	//the times of action 
	class Action_judge   Act_times;
 
	for (Domain.numi = 0; Domain.numi < 100000; Domain.numi++) {  //20000 / 50 = 400 s
		//accel part
#ifdef ACCEL
		//lpDisp =  wsMySheet.get_Range(_variant_t((LONG)1), _variant_t((LONG)(2 + i)));
		Accele_x = (VARIANT)rgMyRge_accel.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(1)));
		Accele_y = (VARIANT)rgMyRge_accel.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(2)));
		Accele_z = (VARIANT)rgMyRge_accel.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(3)));

		bMsg_x = (_bstr_t)Accele_x;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Accel.Daccel_x = _ttof(str_x);

		if (str_x == _T(""))
			break;


		bMsg_x = (_bstr_t)Accele_y;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Accel.Daccel_y = _ttof(str_x);
		if (str_x == _T(""))
			break;


		bMsg_x = (_bstr_t)Accele_z;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Accel.Daccel_z = _ttof(str_x);
		if (str_x == _T(""))
			break;

#endif		
#ifdef GYRO
		//gyro part
		//lpDisp =  wsMySheet.get_Range(_variant_t((LONG)1), _variant_t((LONG)(2 + i)));
		Gyro_x = (VARIANT)rgMyRge_gyro.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(1)));
		Gyro_y = (VARIANT)rgMyRge_gyro.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(2)));
		Gyro_z = (VARIANT)rgMyRge_gyro.get_Item(_variant_t((LONG)2 + Domain.numi), _variant_t((LONG)(3)));

		
		bMsg_x = (_bstr_t)Gyro_x;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Gyro.Dgyro_x = _ttof(str_x);

		if (str_x == _T(""))
			break;


		bMsg_x = (_bstr_t)Gyro_y;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Gyro.Dgyro_y = _ttof(str_x);
		if (str_x == _T(""))
			break;


		bMsg_x = (_bstr_t)Gyro_z;
		strcpy_s(cMsg_x, (const char *)bMsg_x);
		str_x = cMsg_x;
		Gyro.Dgyro_z = _ttof(str_x);
		if (str_x == _T(""))
			break;
		

		//Calculating the Gyro angle
		Get_gyro_angle(Gyro, Accel);
#endif


		//Calculating the accel angle
		Get_accel_angle(Accel);

		

		//G:\\My thesis\\accelerometer_10_times_shakes.xlsx
		//G:\\My thesis\\Gyro_10_times_shakes.xlsx

		//=OFFSET($A$2,COUNTA($A$1:$A$3382)-ROW(A2),)

		//G:\\My thesis\\action2 5 times accele.xlsx
		//G:\\My thesis\\action2 5 times gyro.xlsx
		// G:\\My thesis\\Excel Data Analysis\\Action01\\A1_20_accel.xlsx

		//   G:\\My thesis\\Excel Data Analysis\\Action02\\A2_Accel_20.xlsx

		//G:\\My thesis\\Excel Data Analysis\\GY Analysis\\A2_20\\A2_Accel_20.xlsx

		//G:\\My thesis\\Excel Data Analysis\\GY Analysis\\A1_20\\A1_Accel_20.xlsx
		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\GY Analysis\\A1_20\\A1_Accel_20.xlsx
		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\GY Analysis\\A2_20\\A2_Accel_20.xlsx


		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\Gesture Collections\\Action1\\Gyro_A1_20.xlsx
		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\Gesture Collections\\Action1\\Accel_A1_20.xlsx

		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\Gesture Collections\\Action2\\Accel_A2_20.xlsx

		//F:\\Total Courses\\Thesis\\Excel Data Analysis\\Stand up the phone for understanding the angle change\\A2_10_A1_10_A2_10_A1_10.xlsx

		int act = Action_Calclulation(Accel, Gyro);
		if (act == 1)
			Act_times.Total_a1++;
		else if (act == 2)
			Act_times.Total_a2++;
#ifdef GESTURE_THREE
		else if (act == 3)
			Act_times.Total_a3++;
#endif
		
		str.Format(_T("%d"), Act_times.Total_a1);
		SetDlgItemText(IDC_EDIT3, str);

		str.Format(_T("%d"), Act_times.Total_a2);
		SetDlgItemText(IDC_EDIT4, str);
		
#ifdef GESTURE_THREE		
		str.Format(_T("%d"), Act_times.Total_a3);
		SetDlgItemText(IDC_EDIT5, str);
#endif
	}

	/*
	//accelerometer
CWorkbooks wbsMyBooks_accel;
CWorkbook wbMyBook_accel;
CWorksheets wsMySheets_accel;
CWorksheet wsMySheet_accel;
CRange rgMyRge_accel;
//gyroscope
CWorkbooks wbsMyBooks_gyro;
CWorkbook wbMyBook_gyro;
CWorksheets wsMySheets_gyro;
CWorksheet wsMySheet_gyro;
CRange rgMyRge_gyro;
	*/

	//�ͷ���Դ  

	rgMyRge_accel.ReleaseDispatch();
	wsMySheet_accel.ReleaseDispatch();
	wsMySheets_accel.ReleaseDispatch();
	wbMyBook_accel.ReleaseDispatch();
	wbsMyBooks_accel.ReleaseDispatch();

	rgMyRge_gyro.ReleaseDispatch();
	wsMySheet_gyro.ReleaseDispatch();
	wsMySheets_gyro.ReleaseDispatch();
	wbMyBook_gyro.ReleaseDispatch();
	wbsMyBooks_gyro.ReleaseDispatch();

	ExcelApp.ReleaseDispatch();
	ExcelApp.Quit();
}
