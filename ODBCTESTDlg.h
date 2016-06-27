// ODBCTESTDlg.h : header file
//

#if !defined(AFX_ODBCTESTDLG_H__BF1385C9_AF56_4D9B_A1D1_7406A359DE19__INCLUDED_)
#define AFX_ODBCTESTDLG_H__BF1385C9_AF56_4D9B_A1D1_7406A359DE19__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CODBCTESTDlg dialog

//�ṹ�壬�����ʱ��ʹ��
struct   CompareFunc_lParamSort 
{
	CListCtrl * pl;//ָ��ListView�ؼ���ָ�� 
    int idCol;     //ָ��ListView���� 
    BOOL isInc;    //����ʽ(����ΪTRUE,����ΪFALSE) 
    BOOL isStr;    //�Ƿ��ַ����� 
};

  
class CODBCTESTDlg : public CDialog
{
// Construction
public:
	int count;
	int m_jiangxuejin;
	CODBCTESTDlg(CWnd* pParent = NULL);	// standard constructor

///////////////////////////////////////////////////////////////////////
	//����Զ��庯��
	//��������
	CString GetExcelDriver();    //���excel��������2011-12-6 9:59
	//�ع�����,��������
    static int CALLBACK SortFunc(LPARAM lParam1,LPARAM lParam2,LPARAM lParam3);   

// Dialog Data
	//{{AFX_DATA(CODBCTESTDlg)
	enum { IDD = IDD_ODBCTEST_DIALOG };
	CListCtrl	m_ListCtrl;
	CString	m_name;
	CString	m_college;
	CString	m_age;
	CString	m_salary;
	CString	m_jisuanji;
	CString	m_shuzi;
	CString	m_english;
	float	m_sd;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CODBCTESTDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CODBCTESTDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnImportexcel();
	afx_msg void OnAdd();
	virtual void OnOK();
	afx_msg void OnClickList(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnUpdate();
	afx_msg void OnDelete();
	afx_msg void OnSubmit();
	afx_msg void OnColumnclickList(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnExportexcel();
	afx_msg void OnButton1();
	afx_msg void OnButton2();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ODBCTESTDLG_H__BF1385C9_AF56_4D9B_A1D1_7406A359DE19__INCLUDED_)
