using System;
using System.Windows.Forms;
using System.Threading;

namespace NoSleep
{
    static class NoSleep
    {
        // ミューテックスのインスタンス
        private static Mutex mutex = null;

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // ミューテックス名を設定 (ユニークな名前を付ける)
            string mutexName = "ScreensaverDeterApp";

            // ミューテックスを作成または取得
            mutex = new Mutex(true, mutexName, out bool isNewInstance);

            if (!isNewInstance)
            {
                // 既にアプリケーションが起動している場合
                MessageBox.Show("既に起動しています。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // プロセスを終了
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // フォームを表示しない
            new NoSleepForm();
            Application.Run();

            // アプリ終了時にミューテックスを解放
            GC.KeepAlive(mutex);
        }
    }
}
