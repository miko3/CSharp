/// <summary>
/// PowerShell実行用
/// </summary>
namespace ExecPowerShell
{
    using System;
    using System.Collections.ObjectModel;
    using System.Management.Automation;
    using System.Management.Automation.Runspaces;

    /// <summary>
    /// PowerShell実行用クラス
    /// </summary>
    public class ExecPS
    {
        private Collection<PSObject> adapts;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExecPS()
        {
            this.adapts = null;
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
