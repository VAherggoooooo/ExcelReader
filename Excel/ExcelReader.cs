using System.Collections.Generic;
using UnityEngine;
using Excel;
using System.IO;
using UnityEditor;
using System.Data;

public class ExcelReader : EditorWindow
{
    #region members
    public static Object obj;//物体
    public static readonly string excelExten = ".xlsx";//后缀
    public int tabelID = 0;//表序号
    public List<int> colID = new List<int>();//所需要的列序号
    public List<int> rowID = new List<int>();//所跳过的行序号
    public int startRow, endRow;//起始与终止行
    private Vector2 scrollPos;//滚动界面rect
    public bool showCol = true;//折叠 列
    public bool showRow = true;//折叠 行
    public bool needSE = false;//是否需要手动输入起终行
    public bool showSE = true;//
    #endregion

    #region create window example
    //[MenuItem("Excel/Excel Reader")]
    // public static void ShowWindow()
    // {
    //     GetWindow<ExcelReader>().Show();
    // }
    #endregion

    public virtual void OnGUI()
    {
        #region scrollview
        scrollPos = GUILayout.BeginScrollView(scrollPos);
        {
            #region select excel
            GUILayout.Space(10);
            GUILayout.BeginHorizontal("Box");
            {
                GUILayout.Label("Excel: ");
                Object temp_obj = EditorGUILayout.ObjectField(obj, typeof(Object), true);
                if (temp_obj != null)
                {
                    string path = AssetDatabase.GetAssetPath(temp_obj);
                    string extension = Path.GetExtension(path);
                    if (extension == excelExten)
                    {
                        obj = temp_obj;
                    }
                    else
                    {
                        Debug.LogError("请选择excel文件");
                    }
                }
            }
            GUILayout.EndHorizontal();
            GUILayout.Space(10);
            #endregion

            #region table id
            GUILayout.BeginHorizontal("Box");
            {
                GUILayout.Label("表格序号: ");
                tabelID = EditorGUILayout.IntField(tabelID, GUILayout.MaxWidth(40));
            }
            GUILayout.EndHorizontal();
            GUILayout.Space(10);

            #endregion

            #region start and end row
            GUILayout.BeginVertical("Box");
            {
                needSE = GUILayout.Toggle(needSE, "使用输入的起始终止行号下标");
                if (needSE)
                {
                    GUILayout.Space(5);
                    showSE = EditorGUILayout.Foldout(showSE, "起始终止行下标: ");
                    if (showSE)
                    {
                        GUILayout.BeginVertical("Box");
                        {
                            GUILayout.BeginHorizontal();
                            {
                                GUILayout.Label("起始行下标: ");
                                startRow = EditorGUILayout.IntField(startRow, GUILayout.MaxWidth(40));
                            }
                            GUILayout.EndHorizontal();
                            GUILayout.Space(5);
                            GUILayout.BeginHorizontal();
                            {
                                GUILayout.Label("终止行下标: ");
                                endRow = EditorGUILayout.IntField(endRow, GUILayout.MaxWidth(40));
                            }
                            GUILayout.EndHorizontal();
                        }
                        GUILayout.EndVertical();
                    }
                }
            }
            GUILayout.EndVertical();
            GUILayout.Space(15);
            #endregion

            #region col
            GUILayout.BeginVertical("Box");
            {
                showCol = EditorGUILayout.Foldout(showCol, "需要读取的列序号下标: ");
                if (showCol)
                {
                    GUILayout.BeginVertical("Box");
                    {
                        for (int i = 0; i < colID.Count; i++)
                        {
                            GUILayout.BeginHorizontal("Box");
                            {
                                GUILayout.Label("Col" + i + ": ");
                                colID[i] = EditorGUILayout.IntField(colID[i], GUILayout.MaxWidth(40));
                            }
                            GUILayout.EndHorizontal();
                        }
                        GUILayout.EndVertical();
                    }


                    GUILayout.BeginHorizontal();
                    {
                        GUILayout.Label("");
                        GUILayout.FlexibleSpace();
                        GUILayout.BeginHorizontal("Box");
                        {
                            if (GUILayout.Button("+", GUILayout.MaxWidth(20)))
                            {
                                colID.Add(0);
                            }
                            if (GUILayout.Button("-", GUILayout.MaxWidth(20)))
                            {
                                if (colID.Count > 0)
                                {
                                    colID.RemoveAt(colID.Count - 1);
                                }
                            }
                        }
                        GUILayout.EndHorizontal();
                    }
                    GUILayout.EndHorizontal();
                }
            }
            GUILayout.EndVertical();
            GUILayout.Space(10);
            #endregion

            #region row
            GUILayout.BeginVertical("Box");
            {
                showRow = EditorGUILayout.Foldout(showRow, "需要跳过的行序号下标: ");
                if (showRow)
                {
                    GUILayout.BeginVertical("Box");
                    {
                        for (int i = 0; i < rowID.Count; i++)
                        {
                            GUILayout.BeginHorizontal("Box");
                            {
                                GUILayout.Label("Row" + i + ": ");
                                rowID[i] = EditorGUILayout.IntField(rowID[i], GUILayout.MaxWidth(40));
                            }
                            GUILayout.EndHorizontal();
                        }
                    }
                    GUILayout.EndVertical();

                    GUILayout.BeginHorizontal();
                    {
                        GUILayout.Label("");
                        GUILayout.FlexibleSpace();
                        GUILayout.BeginHorizontal("Box");
                        {
                            if (GUILayout.Button("+", GUILayout.MaxWidth(20)))
                            {
                                rowID.Add(0);
                            }
                            if (GUILayout.Button("-", GUILayout.MaxWidth(20)))
                            {
                                if (rowID.Count > 0)
                                {
                                    rowID.RemoveAt(rowID.Count - 1);
                                }
                            }
                        }
                        GUILayout.EndHorizontal();
                    }
                    GUILayout.EndHorizontal();
                }
            }
            GUILayout.EndVertical();
            GUILayout.Space(10);
            #endregion

            #region custom gui
            GustomGUI();
            GUILayout.Space(10);
            #endregion

            #region button
            GUILayout.BeginHorizontal();
            {
                GUILayout.Label("");
                GUILayout.FlexibleSpace();
                if (GUILayout.Button("Refresh", GUILayout.MaxWidth(90)))
                {
                    obj = null;
                    colID.Clear();
                    rowID.Clear();
                }

                GUI.enabled = (obj == null || colID.Count <= 0) ? false : true;

                if (GUILayout.Button("Read", GUILayout.MaxWidth(90)))
                {
                    ReadExcel(AssetDatabase.GetAssetPath(obj));
                }
            }
            GUILayout.EndHorizontal();
            GUILayout.Space(10);
            #endregion

            #region help
            GUI.enabled = true;
            string msg = string.Empty;
            msg += "使用步骤: \n";
            msg += "·选择excel\n";
            msg += "·输入表格序号(0起)\n";
            msg += "·添加所需列号(0起)\n";
            msg += "·添加跳过行号(0起)\n";
            msg += "·点击read按钮\n";
            msg += "·(refresh按钮为重置当前面板)\n";
            EditorGUILayout.HelpBox(msg, MessageType.Info);
            GUILayout.Space(10);
            #endregion
        }
        GUILayout.EndScrollView();
        #endregion
    }

    #region read function
    public void ReadExcel(string path)
    {
        try
        {
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            DataSet dataSet = excelReader.AsDataSet();
            if (tabelID < 0 || tabelID > dataSet.Tables.Count - 1)
            {
                Debug.LogError("不存在该序号下标的表格");
                return;
            }
            DataTable data = dataSet.Tables[tabelID];
            int line = data.Rows.Count;//行数

            OnReadBegin();

            #region backup
            //int curLine = 0;//当前行号
            // while (excelReader.Read())
            // {
            //     if(!CheckLine(ref curLine)) continue;
            //     string info = string.Empty;
            //     for (int i = 0; i < rowID.Count; i++)//遍历需要的列
            //     {
            //         info += reader.GetString(rowID[i]) + ": ";//将每列内容字符串拼接
            //     }
            //     if (info != string.Empty)
            //     {
            //         Debug.Log(info);//不为空则输出字符串
            //     }
            // }
            #endregion

            int s = needSE ? startRow : 0;
            int e = needSE ? endRow + 1 : line;
            for (int i = s; i < e; i++)//遍历行
            {
                if (!CheckLine(i)) continue;//是否跳过行
                OnReadLine(data, i);
            }

            OnReadEnd();

        }
        catch (IOException)
        {
            Debug.LogError("请关闭表格后读取");
        }
    }
    #endregion

    /// <summary>
    /// 读取开始前
    /// </summary>
    public virtual void OnReadBegin() { }

    /// <summary>
    /// 对于读取的每行执行的内容
    /// </summary>
    /// <param name="data"></param>
    /// <param name="row">当前行号</param>
    public virtual void OnReadLine(DataTable data, int row)
    {
        string info = string.Empty;
        for (int j = 0; j < colID.Count; j++)//遍历列
        {
            if (colID[j] > data.Columns.Count - 1) continue;//超过列数则忽略
            info += data.Rows[row][colID[j]].ToString() + ": ";//将内容字符拼接
        }
        if (info != string.Empty)
        {
            Debug.Log(info);
        }
    }

    /// <summary>
    /// 读取结束后
    /// </summary>
    public virtual void OnReadEnd() { }

    #region custom gui
    public virtual void GustomGUI() { }
    #endregion

    #region check line
    /// <summary>
    /// 检查跳过行
    /// </summary>
    /// <param name="id">当前行序号</param>
    /// <returns></returns>
    private bool CheckLine(int id)
    {
        for (int i = 0; i < rowID.Count; i++)
        {
            if (id == rowID[i])
            {
                return false;
            }
        }
        return true;
    }
    #endregion
}
