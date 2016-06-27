// ODBCTESTDlg.cpp : implementation file
//
#include   <afxpriv.h>
#include "stdafx.h"
#include "ODBCTEST.h"
#include "ODBCTESTDlg.h"
#include <string>
/////////////////////////////////
//���ͷ�ļ�
#include <afxdb.h> 
#include <odbcinst.h>
/////////////////////////////////

/////////////////////////////////
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About


class CAboutDlg : public CDialog
{
public:
	CAboutDlg();
    
// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CODBCTESTDlg dialog

CODBCTESTDlg::CODBCTESTDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CODBCTESTDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CODBCTESTDlg)
	m_name = _T("");
	m_college = _T("");
	m_age = _T("");
	m_salary = _T("");
	m_jisuanji = _T("");
	m_shuzi = _T("");
	m_english = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CODBCTESTDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CODBCTESTDlg)
	DDX_Control(pDX, IDC_LIST, m_ListCtrl);
	DDX_Text(pDX, IDC_NAME, m_name);
	DDX_Text(pDX, IDC_COLLEGE, m_college);
	DDX_Text(pDX, IDC_AGE, m_age);
	DDX_Text(pDX, IDC_SALARY, m_salary);
	DDX_Text(pDX, IDC_NAME2, m_jisuanji);
	DDX_Text(pDX, IDC_AGE2, m_shuzi);
	DDX_Text(pDX, IDC_AGE3, m_english);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CODBCTESTDlg, CDialog)
	//{{AFX_MSG_MAP(CODBCTESTDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_IMPORTEXCEL, OnImportexcel)
	ON_BN_CLICKED(IDC_ADD, OnAdd)
	ON_NOTIFY(NM_CLICK, IDC_LIST, OnClickList)
	ON_BN_CLICKED(IDC_UPDATE, OnUpdate)
	ON_BN_CLICKED(IDC_DELETE, OnDelete)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST, OnColumnclickList)
	ON_BN_CLICKED(IDC_EXPORTEXCEL, OnExportexcel)
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnButton2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CODBCTESTDlg message handlers

BOOL CODBCTESTDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}


	SetIcon(m_hIcon, TRUE);			
	SetIcon(m_hIcon, FALSE);		
	
	// TODO: Add extra initialization here
	
//��ʼ��listcontrol,��ʾ��Ӧ�������������еĿ��
	////////////////////////////////////////////////////////////////////////
    ListView_SetExtendedListViewStyle(m_ListCtrl.m_hWnd, LVS_EX_FULLROWSELECT  | LVS_EX_HEADERDRAGDROP |LVS_EX_GRIDLINES);
	LVCOLUMN column;
	column.mask=LVCF_FMT|LVCF_WIDTH|LVCF_TEXT|LVCF_SUBITEM;
	column.fmt=LVCFMT_LEFT;
	//column.cx=70;

	column.iSubItem=0;
	column.cx=40;
 	column.pszText="ѧ��";
 	this->m_ListCtrl.InsertColumn(0,&column);

 	column.iSubItem=1;
	column.cx=40;
 	column.pszText="����";
	this->m_ListCtrl.InsertColumn(1,&column);

 	column.iSubItem=2;
	column.cx=40;
 	column.pszText="�ļ�";
 	this->m_ListCtrl.InsertColumn(2,&column);

 	column.iSubItem=3;
	column.cx=60;
 	column.pszText="����";
 	this->m_ListCtrl.InsertColumn(3,&column);

	column.iSubItem=4;
	column.cx=60;
 	column.pszText="�����";
 	this->m_ListCtrl.InsertColumn(4,&column);
	
	
	column.iSubItem=5;
	column.cx=80;
 	column.pszText="�����߼�";
 	this->m_ListCtrl.InsertColumn(5,&column);
	column.iSubItem=6;
	column.cx=50;
 	column.pszText="Ӣ��";
 	this->m_ListCtrl.InsertColumn(6,&column);
	
	column.iSubItem=7;
	column.cx=50;
 	column.pszText="ѧ�ּ�";
 	this->m_ListCtrl.InsertColumn(7,&column);
	column.iSubItem=8;
	column.cx=65;
 	column.pszText="��ѧ��";
 	this->m_ListCtrl.InsertColumn(8,&column);
	
    ////////////////////////////////////////////////////////////////////////




	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CODBCTESTDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}




void CODBCTESTDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); 

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CODBCTESTDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CODBCTESTDlg::OnImportexcel() 
{
	// TODO: Add your control notification handler code here
	if(MessageBox("��ȷ���Ѿ���Ҫ�����excel�Ĺ���������Ϊ: ��Ϣͳ��\n����,�޷���������","����",MB_ICONEXCLAMATION|MB_OKCANCEL)==IDCANCEL)return;
		 
	_TCHAR strFileFilter[] = "Microsoft Excel 97-2003 ������(*.xls)|*.xls|"
	    	                 "Microsoft Excel   2007  ������(*.xlsx)|*.xlsx|";
			//"�����ļ�(*.*)|*.*||";
	CFileDialog Dlg(TRUE,NULL ,NULL, OFN_HIDEREADONLY,strFileFilter);
    if(Dlg.DoModal()==IDOK)
	{
	CString strFileName = Dlg.GetFileName();   //��ȡ�ļ���
    CString strFilePath = Dlg.GetPathName();   //��ȡ�ļ�·��
	GetDlgItem(IDC_FILEPATH)->SetWindowText(strFilePath);    //��ʾ�ļ�·��
   

//////////////////////////////////////////////////////////////////////////////////
	CDatabase database;
    CString sSql;
    CString sItem1[10000], sItem2[10000],sItem3[10000],sItem4[10000],sItem5[10000],sItem6[10000],sItem7[10000],sItem8[10000];
    CString sDriver;
    CString sDsn;
	
    // �����Ƿ�װ��Excel���� 
    sDriver = GetExcelDriver();
    if (sDriver.IsEmpty())
    {
        // û�з���Excel����
        AfxMessageBox("û�а�װExcel����!");
        return;
    }
    
    // �������д�ȡ���ַ���
    sDsn.Format("ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver, strFilePath);

    

    TRY
    {
        // �����ݿ�(��Excel�ļ�)
        database.Open(NULL, false, false, sDsn);
        CRecordset recset(&database);
        // ���ö�ȡ�Ĳ�ѯ���.
		
        sSql = "SELECT * " 
			    
               "FROM [��Ϣͳ��$] " ;//�˴��ǹؼ�֮��,��Ҫ��[$],������ʾ�Ҳ����������� 
		       
        // ִ�в�ѯ���
        recset.Open(CRecordset::forwardOnly, sSql, CRecordset::readOnly);

	    m_ListCtrl.DeleteAllItems();    //�˾������������ȫ����¼

		int i=0;
		// ��ȡ��ѯ���
        while (!recset.IsEOF())
        {
            //��ȡExcel�ڲ���ֵ
            recset.GetFieldValue("ѧ��", sItem1[i]);
            recset.GetFieldValue("����", sItem2[i]);
            recset.GetFieldValue("�ļ�", sItem3[i]);
			recset.GetFieldValue("����", sItem4[i]);
			recset.GetFieldValue("�����", sItem5[i]);
			recset.GetFieldValue("�����߼�",sItem6[i]);
			recset.GetFieldValue("Ӣ��", sItem7[i]);
			//��ʾ��ȡ������		
////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////
 
	                 	m_ListCtrl.InsertItem(0, " ");               //**************������������ʾ************************
                       m_ListCtrl.SetItemText(0,0, sItem1[i]); 
                       m_ListCtrl.SetItemText(0,1, sItem2[i]); 
                       m_ListCtrl.SetItemText(0,2, sItem3[i]); 
                       m_ListCtrl.SetItemText(0,3, sItem4[i]);  
					   m_ListCtrl.SetItemText(0,4, sItem5[i]);
					   m_ListCtrl.SetItemText(0,5, sItem6[i]);
					   m_ListCtrl.SetItemText(0,6, sItem7[i]);
					   m_ListCtrl.SetItemText(0,7, sItem8[i]);
				      
            // �Ƶ�excel����һ��
           recset.MoveNext();
		   i=i+1;
        }
        // �ر����ݿ�
        database.Close();                         
    }
    CATCH(CDBException, e)
    {
        // ���ݿ���������쳣ʱ...
        AfxMessageBox("���ݿ����: " + e->m_strError);
    }
    END_CATCH;
//////////////////////////////////////////////////////////////////////////////////

 }
}



//���excel��������
CString CODBCTESTDlg::GetExcelDriver()
{
    char szBuf[2001];
    WORD cbBufMax=2000;
    WORD cbBufOut;
    char *pszBuf=szBuf;
    CString sDriver;

    // ��ȡ�Ѱ�װ����������(������odbcinst.h��)
    if (!SQLGetInstalledDrivers(szBuf,cbBufMax,&cbBufOut))
        return "";
    
    // �����Ѱ�װ�������Ƿ���Excel...
    do
    {
        if (strstr(pszBuf,"Excel")!=0)
        {
            //���� !
            sDriver=CString(pszBuf);
            break;
        }
        pszBuf=strchr(pszBuf,'\0')+1;
    }
    while (pszBuf[1]!='\0');

    return sDriver;
}

void CODBCTESTDlg::OnAdd() 
{
	// TODO: Add your control notification handler code here	
	UpdateData(true);
    if(m_name!="")
	{
		
		
		
			m_ListCtrl.InsertItem(0, " ");     //������������ʾ
            m_ListCtrl.SetItemText(0,0, m_name);
            m_ListCtrl.SetItemText(0,1, m_age); 
            m_ListCtrl.SetItemText(0,2, m_college); 
            m_ListCtrl.SetItemText(0,3, m_salary);
			m_ListCtrl.SetItemText(0,4, m_jisuanji);
			m_ListCtrl.SetItemText(0,5, m_shuzi);
			m_ListCtrl.SetItemText(0,6, m_english);
			MessageBox("�����ɹ�","��Ϣ��ʾ",64);
		
		    GetDlgItem(IDC_NAME)->SetWindowText("");//����ı���
		    GetDlgItem(IDC_AGE)->SetWindowText("");
		    GetDlgItem(IDC_COLLEGE)->SetWindowText("");
		    GetDlgItem(IDC_SALARY)->SetWindowText("");
			GetDlgItem(IDC_AGE2)->SetWindowText("");
			GetDlgItem(IDC_AGE3)->SetWindowText("");
			GetDlgItem(IDC_NAME2)->SetWindowText("");
		
	}
	else
		MessageBox("ѧ�Ų���Ϊ��","ERROR",16);
}

void CODBCTESTDlg::OnOK() 
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();
}


//��ȡĳһ�еĽ���,����������ʾ���ı���
void CODBCTESTDlg::OnClickList(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
	if(pos)
	{   int x;
		int nItem = m_ListCtrl.GetNextSelectedItem(pos);
		CString strUserinfo;
		strUserinfo.Format("ѧ��:%s ����:%s �ļ�:%s ����:%s �����:%s �����߼�:%s Ӣ��:%s ѧ�ּ�:%s",m_ListCtrl.GetItemText(nItem,0),m_ListCtrl.GetItemText(nItem,1),m_ListCtrl.GetItemText(nItem,2),m_ListCtrl.GetItemText(nItem,3),m_ListCtrl.GetItemText(nItem,4),m_ListCtrl.GetItemText(nItem,5),m_ListCtrl.GetItemText(nItem,6),m_ListCtrl.GetItemText(nItem,7));
		GetDlgItem(IDC_DATA)->SetWindowText(strUserinfo);
		GetDlgItem(IDC_NAME)->SetWindowText(m_ListCtrl.GetItemText(nItem,0));
		GetDlgItem(IDC_AGE)->SetWindowText(m_ListCtrl.GetItemText(nItem,1));
		GetDlgItem(IDC_COLLEGE)->SetWindowText(m_ListCtrl.GetItemText(nItem,2));
		GetDlgItem(IDC_SALARY)->SetWindowText(m_ListCtrl.GetItemText(nItem,3));
		GetDlgItem(IDC_NAME2)->SetWindowText(m_ListCtrl.GetItemText(nItem,4));
		GetDlgItem(IDC_AGE2)->SetWindowText(m_ListCtrl.GetItemText(nItem,5));
	    GetDlgItem(IDC_AGE3)->SetWindowText(m_ListCtrl.GetItemText(nItem,6));
       
	   
	}
	*pResult = 0;
}

void CODBCTESTDlg::OnUpdate() 
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if(m_ListCtrl.GetItemCount()==0)  //ListCtrl��û�����ݿ��Ե���
	{
		MessageBox("û�����ݿ����޸�","��ʾ",48);
		return;
	}
	if(m_ListCtrl.GetNextItem(-1,LVNI_SELECTED)==-1) //û��ѡ��ListCtrl�е�����
	{
		MessageBox("��ѡ��Ҫ�޸ĵ�����","��ʾ",48);
		return;
	}
    if(m_name!="")
	{
		
		
			m_ListCtrl.InsertItem(0, " ");               //**************������������ʾ************************
            m_ListCtrl.SetItemText(0,0, m_name);
            m_ListCtrl.SetItemText(0,1, m_age); 
            m_ListCtrl.SetItemText(0,2, m_college); 
            m_ListCtrl.SetItemText(0,3, m_salary); 
			 m_ListCtrl.SetItemText(0,4, m_jisuanji);
			  m_ListCtrl.SetItemText(0,5, m_shuzi);
			   m_ListCtrl.SetItemText(0,6, m_english);

			m_ListCtrl.DeleteItem(m_ListCtrl.GetNextItem(-1,LVNI_SELECTED));  //�������,����ֻ���޸ĵ�һ����¼ 
            MessageBox("�����ɹ�","��Ϣ��ʾ",64);
		
	}
    else
		MessageBox("��������Ϊ��","ERROR",16);
}

void CODBCTESTDlg::OnDelete() 
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if(m_ListCtrl.GetItemCount()==0)  //ListCtrl��û�����ݿ��Ե���
	{
		MessageBox("û�����ݿ���ɾ��","��ʾ",48);
		return;
	}
	if(m_ListCtrl.GetNextItem(-1,LVNI_SELECTED)==-1) //û��ѡ��ListCtrl�е�����
	{
		MessageBox("��ѡ��Ҫɾ��������","��ʾ",48);
		return;
	}
	if(MessageBox("��ȷ��Ҫɾ��������¼��?\n\n����:"+m_name+"","��Ϣ��ʾ",MB_ICONEXCLAMATION|MB_OKCANCEL)==IDCANCEL)return;
	m_ListCtrl.DeleteItem(m_ListCtrl.GetNextItem(-1,LVNI_SELECTED));  //��仰��֤�˲���ֻ���޸ĵ�һ����¼ 
    
	//����ı���
	GetDlgItem(IDC_NAME)->SetWindowText("");
	GetDlgItem(IDC_AGE)->SetWindowText("");
    GetDlgItem(IDC_COLLEGE)->SetWindowText("");
    GetDlgItem(IDC_SALARY)->SetWindowText("");
	GetDlgItem(IDC_AGE2)->SetWindowText("");
	GetDlgItem(IDC_AGE3)->SetWindowText("");
	GetDlgItem(IDC_NAME2)->SetWindowText("");
}











int CALLBACK CODBCTESTDlg::SortFunc(LPARAM lParam1,LPARAM lParam2,LPARAM lParam3) 
{
	CompareFunc_lParamSort * pparam; 
    CString cstr,cstr2; 
    int ret;   

    pparam=(CompareFunc_lParamSort *)lParam3;     
    cstr=pparam->pl->GetItemText(lParam1,pparam->idCol);     
    cstr2=pparam->pl->GetItemText(lParam2,pparam->idCol);   
    if(pparam->isStr)
		ret=strcmp(cstr,cstr2); 
    else 
        ret=atof(cstr)<=atof(cstr2)?-1:1; 
    if(pparam->isInc) 
        return ret; 
    else 
        return -ret;   
}   


void CODBCTESTDlg::OnColumnclickList(NMHDR* pNMHDR,LRESULT* pResult) //�����ͷ����
{
	NM_LISTVIEW* pLV=(NM_LISTVIEW*)pNMHDR; 

    CompareFunc_lParamSort sortparam; 
    static int idColindex=-1;      //��ǰ�к� 
    static BOOL bInc=TRUE;         //����ʽ 
    int i; 

    if(pLV->iSubItem==idColindex)   //����ǵ�ǰ��,��ı�����ʽ
        bInc=!bInc; 
    else  //������ǵ�ǰ��,�������к�,����Ϊ�������� 
	{
		idColindex=pLV-> iSubItem; 
        bInc=TRUE; 
    } 

    i=m_ListCtrl.GetItemCount(); 
    while(i--)m_ListCtrl.SetItemData(i,i);     //ע��˾�:ItemData���������ʹ�� 
    sortparam.pl=&m_ListCtrl; 
    sortparam.idCol=idColindex; 
    sortparam.isInc=bInc; 
    sortparam.isStr=TRUE; 
    m_ListCtrl.SortItems(SortFunc,(DWORD)&sortparam); 
    *pResult=0; 
}

void CODBCTESTDlg::OnExportexcel() 
{
	// TODO: Add your control notification handler code here
	if(m_ListCtrl.GetItemCount()==0)
	{
		MessageBox("û�����ݿ��Ե���","��ʾ",48);
		return;
	}
    static char BASED_CODE szExt[]="��Ϣͳ��";
	static char BASED_CODE szName[]="ѧ����Ϣ";
	static char BASED_CODE szFilter[]="Microsoft Excel 97-2003 ������(*.xls)|*.xls|"
		                              "Microsoft Excel   2007  ������(*.xlsx)|*.xlsx|";
	
	CFileDialog dlg(FALSE,szExt,szName,OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,szFilter);
	if(dlg.DoModal()!=IDOK)
		return ;

	CString strFile=dlg.GetPathName();

	FILE *pExcel=NULL;
	if( !(pExcel=fopen(strFile.GetBuffer(0),"w")))
	{
		AfxMessageBox("�����ļ�ʧ��");
		return;
	}

	LVCOLUMN col;
	HDITEM item;
	int nCol=m_ListCtrl.GetHeaderCtrl()->GetItemCount();
	for(int i=0;i<nCol;i++)
	{
		memset(&col,0x00,sizeof(LVCOLUMN));
		memset(&item,0x00,sizeof(HDITEM));
		m_ListCtrl.GetColumn(i,&col);
		m_ListCtrl.GetHeaderCtrl()->GetItem(i,&item);

		fprintf(pExcel,"%s\t",col.pszText);
	}
	fprintf(pExcel,"\n");

	int nRols=m_ListCtrl.GetItemCount();
	CString strText;
	for(int j=0;j<nRols;j++)
	{
		for(i=0;i<nCol;i++)
		{
			strText=m_ListCtrl.GetItemText(j,i);
			fprintf(pExcel,"%s\t",strText);
		}

		fprintf(pExcel,"\n");
	}
	fclose(pExcel);
    MessageBox("�������","��ʾ",48);
	
	

}

void CODBCTESTDlg::OnButton1() //ѧ�ּ�����
{MessageBox("ע�����������ѧ����ռѧ�ּ����ٽ��м��㣡","��ʾ",48);
	float p;CString j1,j2,j3,j4;
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
int nItem = m_ListCtrl.GetNextSelectedItem(pos);
GetDlgItemText(IDC_EDIT1,j1);
GetDlgItemText(IDC_EDIT2,j2);
GetDlgItemText(IDC_EDIT3,j3);
GetDlgItemText(IDC_EDIT4,j4);
for(int i=0;i<999;i++)
{float q=atol(m_ListCtrl.GetItemText(nItem,3));
float x=atol(m_ListCtrl.GetItemText(nItem,4));
float j=atol(m_ListCtrl.GetItemText(nItem,5));
float k=atol(m_ListCtrl.GetItemText(nItem,6));
float j5=atol(j1);
float j6=atol(j2);
float j7=atol(j3);
float j8=atol(j4);
p=(q*j5+x*j6+j*j7+k*j8)/(j5+j6+j7+j8);
		CString str;          
          str.Format("%4.2f",p);
		  m_ListCtrl.SetItemText(nItem,7, str);
		  nItem++;
}
		   }


void CODBCTESTDlg::OnButton2() //��ѧ�����
{
	// TODO: Add your control notification handler code here
	
	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
int nItem = m_ListCtrl.GetNextSelectedItem(pos);

for(int i=0;i<999;i++)
{int w=atol(m_ListCtrl.GetItemText(nItem,2));
int q=atol(m_ListCtrl.GetItemText(nItem,3));
int x=atol(m_ListCtrl.GetItemText(nItem,4));
int j=atol(m_ListCtrl.GetItemText(nItem,5));
int k=atol(m_ListCtrl.GetItemText(nItem,6));
int z=atol(m_ListCtrl.GetItemText(nItem,7));

if(w>=355&&z>=75&&q>=70&&x>=70&&j>=70&&k>=70)
{m_ListCtrl.SetItemText(nItem,8, "3-���Ƚ�");

if(w>=355&&z>=82&&q>=75&&x>=75&&j>=75&&k>=75)
{m_ListCtrl.SetItemText(nItem,8, "2-���Ƚ�");


if(w>=355&&z>=88&&q>=80&&x>=80&&j>=80&&k>=80)
{m_ListCtrl.SetItemText(nItem,8, "1-һ�Ƚ�");


if(w>=355&&z>=92&&q>=86&&x>=86&&j>=86&&k>=86)
{m_ListCtrl.SetItemText(nItem,8, "0-�صȽ�");}
}}}
else {m_ListCtrl.SetItemText(nItem,8, "û���");}
		  nItem++;
}
}
