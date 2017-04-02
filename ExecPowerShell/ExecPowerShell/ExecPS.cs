/// <summary>
/// PowerShell実行用
/// </summary>
namespace ExecPowerShell
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Management.Automation;
    using System.Management.Automation.Runspaces;

    /// <summary>
    /// PowerShell実行用クラス
    /// </summary>
    public class ExecPS
    {
        /// <summary>
        /// PowerShell引数のタイプ
        /// </summary>
        public enum paramType{
            ARGMENT,PARAMETER,
        }

        /// <summary>
        /// パラメータ保持用
        /// </summary>
        public struct ParamStruct
        {
            public paramType paramType;
            public string paramName;
            public string paramValue;
        }

        /// <summary>
        /// 引数用
        /// </summary>
        private Collection<PSObject> adapts;

        /// <summary>
        /// パラメータ保持用リスト
        /// </summary>
        private List<ParamStruct> paramList;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExecPS()
        {
            this.adapts = null;
            paramList = new List<ParamStruct>();
        }

        /// <summary>
        /// パラメータを設定
        /// </summary>
        /// <param name="paramList">パラメータ保持用リスト</param>
        /// <param name="pt">パラメータのタイプを指定する(Argment, Parameter)</param>
        /// <param name="paramName"></param>
        /// <param name="paramValue"></param>
        public static void AddParam(ref List<ParamStruct> paramList, paramType pt, string paramName, string paramValue = null)
        {
            try
            {
                ParamStruct pStruct = new ParamStruct();
                // パラメータタイプを設定
                pStruct.paramType = pt;
                // パラメータ名を指定
                pStruct.paramName = paramName;
                // パラメータにセットする値を設定
                pStruct.paramValue = paramValue;

                // リストに追加
                paramList.Add(pStruct);
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }

        /// <summary>
        /// PSスクリプト実行（引数なし）
        /// </summary>
        /// <param name="psScriptName">実行PSスクリプト名</param>
        /// <returns>PS実行結果</returns>
        public Collection<PSObject> ExecPowerShell(string psScriptName)
        {
            try
            {
                using (Runspace rs = RunspaceFactory.CreateRunspace())
                {
                    // Runspace をオープンする
                    rs.Open();

                    using (PowerShell ps = PowerShell.Create())
                    {
                        PSCommand psCmd = new PSCommand();

                        psCmd.AddCommand(psScriptName);

                        ps.Commands = psCmd;
                        ps.Runspace = rs;

                        // スクリプトを実行する
                        adapts = ps.Invoke();
                    }
                }

                return adapts;
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }
 
        /// <summary>
        /// PSスクリプト実行（引数設定済み）
        /// </summary>
        /// <param name="psScriptName"></param>
        /// <returns></returns>
        public Collection<PSObject> ExecPowerShell(string psScriptName, List<ParamStruct> paramList)
        {
            // パラメータがセットされていない場合は終了
            if (paramList == null)
            {
                return null;
            }

            try
            {
                using (Runspace rs = RunspaceFactory.CreateRunspace())
                {
                    // Runspace をオープンする
                    rs.Open();

                    using (PowerShell ps = PowerShell.Create())
                    {
                        PSCommand psCmd = new PSCommand();

                        psCmd.AddCommand(psScriptName);

                        foreach (ParamStruct item in paramList)
                        {
                            switch (item.paramType)
                            {
                                case paramType.ARGMENT:
                                    psCmd.AddArgument(item.paramName);
                                    break;
                                case paramType.PARAMETER:
                                    if (item.paramValue == null)
                                    {
                                        psCmd.AddParameter(item.paramName);
                                    }
                                    else
                                    {
                                        psCmd.AddParameter(item.paramName, item.paramValue);
                                    }
                                    break;
                                default:
                                    return null;
                            }
                        }

                        ps.Commands = psCmd;
                        ps.Runspace = rs;

                        // スクリプトを実行する
                        adapts = ps.Invoke();
                    }
                }

                return adapts;
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }

        /// <summary>
        /// PSスクリプト実行（引数あり）
        /// </summary>
        /// <param name="psScriptName">実行スクリプト名</param>
        /// <param name="psdc">引数</param>
        /// <returns>PS実行結果</returns>
        public Collection<PSObject> ExecPowerShell<T>(string psScriptName, PSDataCollection<T> psdc)
        {
            try
            {
                using (Runspace rs = RunspaceFactory.CreateRunspace())
                {
                    // Runspace をオープンする
                    rs.Open();

                    using (PowerShell ps = PowerShell.Create())
                    {
                        PSCommand psCmd = new PSCommand();

                        psCmd.AddCommand(psScriptName);

                        ps.Commands = psCmd;
                        ps.Runspace = rs;


                        // スクリプトを実行する(配列を渡す)
                        if (psdc != null)
                        {
                            adapts = ps.Invoke(psdc);
                        }
                        else
                        {
                            adapts = ps.Invoke();
                        }
                    }
                }

                return adapts;
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }
    }
}
