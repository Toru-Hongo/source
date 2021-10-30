using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TmWatch {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void tmWatch_Tick(System.Object sender, System.EventArgs e) {
            long sLp;
            DataSet DBSet2 = new DataSet();
            DataSet DBSet3 = new DataSet();
            DataTable DBTable3;
            DataRow row;
            DataRow row2;
            DataTable DBTable;
            double lngValueBuf;
            long lngValueDiff;
            string[] strBuf;
            string strToumei, strHeyamei, strDate;
            short sPLCNo, sPageNo, sRow, sCol;
            short rinf;
            short sModeFlg;
            string strTime;
            double lngGeppouBuf;
            string strFinDateBuf;
            string strPath;
            DateTime dtWorkDate;
            string strDayBuf;
            // 日付変換処理用
            string strWorkBuf;
            string strDay;
            DateTime dtDate;


            Application.DoEvents();

            lblTime.Text = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH時mm分");

            try {

                // 時替り処理
                if (gTimeBuf != Strings.Format(DateTime.Now, "%H")) {
                    
                    for (sLp = 0; sLp <= 3000; sLp++) {
                        lngValueBuf = 0;
                        if (AlarmData(sLp).mToumei == "" & AlarmData(sLp).mHeyamei == "")
                            // 共にＮＵＬＬなら処理終了
                            break;


                        #region 検索件数が０以上なら、アラーム処理を行う　メータ値
                        // データ取得
                        // ＳＱＬ文作成
                        strSQL = "SELECT メータ値 FROM Ｄ瞬時データ WHERE 棟名称 = '" + AlarmData(sLp).mToumei + "' AND 部屋名称 = '" + AlarmData(sLp).mHeyamei + "'";
                        // ＤＢ読込み
                        DBTable = DBClass.Read2(connectionstring, DBSet, strSQL);

                        // 検索件数が０以上なら、アラーム処理を行う
                        if (DBTable.Rows.Count > 0) {
                            foreach (var row in DBTable.Rows)
                                lngValueBuf = Format(row("メータ値"), "####0.#");

                            // データテーブル初期化
                            DBTable.Clear();

                            if (gintINITPFFlg(sLp) == 1) {
                                AlarmData(sLp).mHourBuf = lngValueBuf;

                                gintINITPFFlg(sLp) = 0;
                            }
                            else {
                                if (lngValueBuf < AlarmData(sLp).mHourBuf) {
                                    // 過大消費電力（警報）
                                    // ＷＨＭとＭＯＦで乗率が違う為、種別を取得する
                                    strSQL = "SELECT シャフト番号 FROM Ｍ登録マスタ WHERE 棟名称='" + AlarmData(sLp).mToumei + "' And 部屋名称='" + AlarmData(sLp).mHeyamei + "'";
                                    // ＤＢ読込み
                                    DBTable3 = DBClass.Read2(connectionstring, DBSet3, strSQL);

                                    if (DBTable3.Rows.Count > 0) {
                                        long lngdiff2;
                                        row2 = DBTable3.Rows(0);

                                        if (row2("シャフト番号") < 17) {

                                            // 前回値より今回値が小さい場合（カウンタが一周した）
                                            // (99999.9 - 前回値) + 今回値=差分
                                            lngdiff2 = 99999.9 - AlarmData(sLp).mHourBuf;
                                            lngValueDiff = lngValueBuf - lngdiff2;
                                        }
                                        else {
                                            // 前回値より今回値が小さい場合（カウンタが一周した）
                                            // (999999(MOF) - 前回値) + 今回値=差分
                                            lngdiff2 = 999999 - AlarmData(sLp).mHourBuf;
                                            lngValueDiff = lngValueBuf - lngdiff2;
                                        }

                                        DBTable3.Clear();
                                    }
                                }
                                else
                                    lngValueDiff = lngValueBuf - AlarmData(sLp).mHourBuf;

                                // '2007-05-17 Change >>>
                                // '戸別とMOFの警報基準値を分ける為、変更
                                // '種別が2以外の時は、戸別として計算
                                // '（発電機も戸別扱いとする）
                                // 2008-07-14 Change >>>
                                // 個別と発電機の警報基準値を分ける為、変更
                                // 種別０は戸別、２はＭＯＦ、３は発電機（１は共用部なので未使用）とする。
                                if (AlarmData(sLp).mSyubetu == 0) {
                                    // 戸別
                                    if (lngValueDiff > mINIData.mExcPower) {
                                        // ワーニング発生
                                        // 表示位置の特定
                                        strBuf = Split(AlarmData(sLp).mArrayPos, ",");

                                        // ＰＬＣ番号
                                        sPLCNo = System.Convert.ToInt16(strBuf[0]);
                                        // 頁番号
                                        sPageNo = System.Convert.ToInt16(strBuf[1]);
                                        // 行
                                        sRow = System.Convert.ToInt16(strBuf[2]);
                                        // 列
                                        sCol = System.Convert.ToInt16(strBuf[3]);

                                        if (gAlarmFlg == 0 & gMeterMode != "1") {
                                            // 他の警報処理で送信中でなければ、送信する
                                            if (sSending == 0) {
                                                // POP3サーバ認証に時間が掛かった場合に、警報メールが複数出てしまうので、
                                                // 送信中は、メールを送らないようにフラグを追加する
                                                sSending = 1;
                                                // Ｅメール発信(既に発生済みでなく、Ｅメール版で且つ、メータ交換モード中でない場合）
                                                strWorkBuf = "棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。";
                                                rinf = Send_Mail(0, strWorkBuf);
                                                // 2008-06-06 T.Isano Add >>>
                                                // 送信結果判断処理
                                                if (rinf == -1) {
                                                    gintRetFlg1 = 1;

                                                    // リトライカウンタ初期化
                                                    gintRetCnt = 0;

                                                    // 警報メール本文再作成
                                                    gstrALMailMessage = "【リトライ送信】：初回メール失敗日時 " + Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + Constants.vbCrLf + strWorkBuf;

                                                    // リトライ用タイマが起動していなければ、起動
                                                    if (tmRetry.Enabled == false)
                                                        tmRetry.Enabled = true;
                                                }
                                                // <<< Add End

                                                sSending = 0;
                                            }
                                        }

                                        // 警報履歴処理作成(正常又は警報状態なら作成)
                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            rinf = Create_AlarmSamry(AlarmData(sLp).mToumei, AlarmData(sLp).mHeyamei, 1);
                                            // 2007-04-11 Add >>>
                                            tmBuzzer.Enabled = true;
                                            // <<< Add ENd

                                            // 2007-05-15 For Debug >>>
                                            string strErr2;
                                            // ログ出力文字列作成
                                            strErr2 = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + ":棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。" + Constants.vbCrLf;
                                            // ログ出力
                                            rinf = Log_Make(strErr2);
                                        }

                                        // 'ブザー鳴動（ブザータイマ開始)
                                        // If gAlarmFlg = 0 Then
                                        // tmBuzzer.Enabled = True
                                        // End If

                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            // 復旧は異常確認ボタンで行う
                                            ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) = 1;
                                            if (ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) < 3) {
                                                ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) = 2;
                                                ViewData(sPLCNo).mFlick(sPageNo, sRow, sCol) = 1;
                                            }
                                        }

                                        gAlarmFlg = 1;
                                    }

                                    AlarmData(sLp).mHourBuf = lngValueBuf;
                                }
                                else if (AlarmData(sLp).mSyubetu == 2) {
                                    // ＭＯＦ
                                    if (lngValueDiff > mINIData.mMOFExcPower) {
                                        // ワーニング発生
                                        // 表示位置の特定
                                        strBuf = Split(AlarmData(sLp).mArrayPos, ",");

                                        // ＰＬＣ番号
                                        sPLCNo = System.Convert.ToInt16(strBuf[0]);
                                        // 頁番号
                                        sPageNo = System.Convert.ToInt16(strBuf[1]);
                                        // 行
                                        sRow = System.Convert.ToInt16(strBuf[2]);
                                        // 列
                                        sCol = System.Convert.ToInt16(strBuf[3]);

                                        if (gAlarmFlg == 0 & gMeterMode != "1") {
                                            if (sSending == 0) {
                                                // POP3サーバ認証に時間が掛かった場合に、警報メールが複数出てしまうので、
                                                // 送信中は、メールを送らないようにフラグを追加する
                                                sSending = 1;
                                                // Ｅメール発信(既に発生済みでなく、Ｅメール版で且つ、メータ交換モード中でない場合）
                                                strWorkBuf = "棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。";
                                                rinf = Send_Mail(0, strWorkBuf);

                                                // 2008-06-06 T.Isano Add >>>
                                                // 送信結果判断処理
                                                if (rinf == -1) {
                                                    gintRetFlg1 = 1;

                                                    // リトライカウンタ初期化
                                                    gintRetCnt = 0;

                                                    // 警報メール本文再作成
                                                    gstrALMailMessage = "【リトライ送信】：初回メール失敗日時 " + Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + Constants.vbCrLf + strWorkBuf;

                                                    // リトライ用タイマが起動していなければ、起動
                                                    if (tmRetry.Enabled == false)
                                                        tmRetry.Enabled = true;
                                                }
                                                // <<< Add End

                                                sSending = 0;
                                            }
                                        }

                                        // 警報履歴処理作成(正常又は警報状態なら作成)
                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            rinf = Create_AlarmSamry(AlarmData(sLp).mToumei, AlarmData(sLp).mHeyamei, 1);
                                            // 2007-04-11 Add >>>
                                            tmBuzzer.Enabled = true;
                                            // <<< Add ENd

                                            // 2007-05-15 For Debug >>>
                                            string strErr2;
                                            // ログ出力文字列作成
                                            strErr2 = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + ":棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。" + Constants.vbCrLf;
                                            // ログ出力
                                            rinf = Log_Make(strErr2);
                                        }

                                        // 'ブザー鳴動（ブザータイマ開始)
                                        // If gAlarmFlg = 0 Then
                                        // tmBuzzer.Enabled = True
                                        // End If

                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            // 復旧は異常確認ボタンで行う
                                            ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) = 1;
                                            if (ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) < 3) {
                                                ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) = 2;
                                                ViewData(sPLCNo).mFlick(sPageNo, sRow, sCol) = 1;
                                            }
                                        }

                                        gAlarmFlg = 1;
                                    }

                                    AlarmData(sLp).mHourBuf = lngValueBuf;
                                }
                                else if (AlarmData(sLp).mSyubetu == 3) {
                                    // 発電機
                                    if (lngValueDiff > mINIData.mCGSExcPower) {
                                        // ワーニング発生
                                        // 表示位置の特定
                                        strBuf = Split(AlarmData(sLp).mArrayPos, ",");

                                        // ＰＬＣ番号
                                        sPLCNo = System.Convert.ToInt16(strBuf[0]);
                                        // 頁番号
                                        sPageNo = System.Convert.ToInt16(strBuf[1]);
                                        // 行
                                        sRow = System.Convert.ToInt16(strBuf[2]);
                                        // 列
                                        sCol = System.Convert.ToInt16(strBuf[3]);

                                        if (gAlarmFlg == 0 & gMeterMode != "1") {
                                            if (sSending == 0) {
                                                // POP3サーバ認証に時間が掛かった場合に、警報メールが複数出てしまうので、
                                                // 送信中は、メールを送らないようにフラグを追加する
                                                sSending = 1;
                                                // Ｅメール発信(既に発生済みでなく、Ｅメール版で且つ、メータ交換モード中でない場合）
                                                strWorkBuf = "棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。";
                                                rinf = Send_Mail(0, strWorkBuf);

                                                // 2008-06-06 T.Isano Add >>>
                                                // 送信結果判断処理
                                                if (rinf == -1) {
                                                    gintRetFlg1 = 1;

                                                    // リトライカウンタ初期化
                                                    gintRetCnt = 0;

                                                    // 警報メール本文再作成
                                                    gstrALMailMessage = "【リトライ送信】：初回メール失敗日時 " + Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + Constants.vbCrLf + strWorkBuf;

                                                    // リトライ用タイマが起動していなければ、起動
                                                    if (tmRetry.Enabled == false)
                                                        tmRetry.Enabled = true;
                                                }
                                                // <<< Add End

                                                sSending = 0;
                                            }
                                        }

                                        // 警報履歴処理作成(正常又は警報状態なら作成)
                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            rinf = Create_AlarmSamry(AlarmData(sLp).mToumei, AlarmData(sLp).mHeyamei, 1);
                                            // 2007-04-11 Add >>>
                                            tmBuzzer.Enabled = true;
                                            // <<< Add ENd

                                            // 2007-05-15 For Debug >>>
                                            string strErr2;
                                            // ログ出力文字列作成
                                            strErr2 = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + ":棟名称:" + Trim(AlarmData(sLp).mToumei) + ",部屋名称:" + Trim(AlarmData(sLp).mHeyamei) + "の電力計で過大消費電力が発生しました。" + Constants.vbCrLf;
                                            // ログ出力
                                            rinf = Log_Make(strErr2);
                                        }

                                        // 'ブザー鳴動（ブザータイマ開始)
                                        // If gAlarmFlg = 0 Then
                                        // tmBuzzer.Enabled = True
                                        // End If

                                        if (ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) == 0) {
                                            // 復旧は異常確認ボタンで行う
                                            ViewData(sPLCNo).mOutPF(sPageNo, sRow, sCol) = 1;
                                            if (ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) < 3) {
                                                ViewData(sPLCNo).mViewBuf(sPageNo, sRow, sCol) = 2;
                                                ViewData(sPLCNo).mFlick(sPageNo, sRow, sCol) = 1;
                                            }
                                        }

                                        gAlarmFlg = 1;
                                    }

                                    AlarmData(sLp).mHourBuf = lngValueBuf;
                                }
                            }
                        }

                        #endregion


                        #region Ｄ日報データ作成

                        // 日報データ作成
                        // レコード有無の確認
                        strToumei = AlarmData(sLp).mToumei;
                        strHeyamei = AlarmData(sLp).mHeyamei;
                        // 日付の設定
                        dtWorkDate = DateTime.Now;
                        strDate = Strings.Format(dtWorkDate, "yyyy年MM月dd日");
                        strSQL = "SELECT 日報日付 FROM Ｄ日報データ WHERE 日報日付='" + strDate + "' And 棟名称='"
                                    + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                        // ＤＢ読込み
                        DBTable = DBClass.Read2(connectionstring, DBSet2, strSQL);

                        // 検索件数が０以上なら、上書きモード
                        if (DBTable.Rows.Count > 0) {

                            // モードフラグ
                            sModeFlg = 1;

                            // データセットクリア
                            DBTable.Clear();
                        }
                        else
                            sModeFlg = 0;

                        // モードによって、発行するＳＱＬ文を変える
                        if (sModeFlg == 0) {
                            // INSERT(日報日付,棟名称,部屋名称,
                            // ０時メータ値,１時,２時,３時,４時,５時,６時,７時,８時,９時
                            // １０時,１１時,１２時,１３時,１４時,１５時,１６時,１７時,１８時,１９時
                            // ２０時,２１時,２２時,２３時);
                            strSQL = "INSERT INTO Ｄ日報データ VALUES('" + strDate + "','" + strToumei + "','" + strHeyamei + "',"
                                        + "0,0,0,0,0,0,0,0,0,0,"
                                        + "0,0,0,0,0,0,0,0,0,0,"
                                        + "0,0,0,0);";

                            // 書込み
                            rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);

                            // UPDATE
                            // 更新用文字列の作成
                            gTimeBuf = Strings.Format(dtWorkDate, "%H");
                            strTime = "メータ値" + StrConv(gTimeBuf, VbStrConv.Wide) + "時";

                            // lngValueBuf = Int(lngValueBuf)
                            strSQL = "UPDATE Ｄ日報データ SET " + strTime + "='" + Strings.Format(lngValueBuf, "#####0.0") + "' WHERE 日報日付='" + strDate + "' And 棟名称='" + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                            // 書込み
                            rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);
                        }
                        else {
                            // UPDATE
                            // 更新用文字列の作成
                            gTimeBuf = Strings.Format(dtWorkDate, "%H");
                            strTime = "メータ値" + StrConv(gTimeBuf, VbStrConv.Wide) + "時";

                            // lngValueBuf = Int(lngValueBuf)
                            strSQL = "UPDATE Ｄ日報データ SET " + strTime + "='" + Strings.Format(lngValueBuf, "#####0.0") + "' WHERE 日報日付='" + strDate + "' And 棟名称='" + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                            // 書込み
                            rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);
                        }

                        // データテーブル初期化
                        DBTable.Clear();
                    }

                    #endregion

                    // 日替り処理
                    if (gDayBuf != Strings.Format(DateTime.Now, "%d")) {



                        #region Ｄ日報データ
                        for (sLp = 0; sLp <= 3000; sLp++) {
                            if (AlarmData(sLp).mToumei == "" & AlarmData(sLp).mHeyamei == "")
                                // 共にＮＵＬＬなら処理終了
                                break;

                            // 棟名称、部屋名称の取得
                            strToumei = AlarmData(sLp).mToumei;
                            strHeyamei = AlarmData(sLp).mHeyamei;
                            strDate = Strings.Format(dtWorkDate, "yyyy年MM月dd日");
                            strDayBuf = Strings.Format(dtWorkDate, "%d");
                            // strDate = gYearBuf & "年" & gMonBuf & "月" & gDayBuf & "日"

                            // 日報データの取得
                            strSQL = "SELECT メータ値０時 FROM Ｄ日報データ WHERE 日報日付='" + strDate + "' And 棟名称='"
                                        + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                            // ＤＢ読込み
                            DBTable = DBClass.Read2(connectionstring, DBSet2, strSQL);

                            // 検索件数が０以上なら、データ格納
                            if (DBTable.Rows.Count > 0) {
                                row = DBTable.Rows(0);

                                lngGeppouBuf = row("メータ値０時");

                                // データテーブル初期化
                                DBTable.Clear();
                            }
                            else
                                // データなし
                                lngGeppouBuf = 0;

                            // レコード有無の確認
                            strDate = Strings.Format(dtWorkDate, "yyyy年MM月");
                            strSQL = "SELECT 月報日付 FROM Ｄ月報データ WHERE 月報日付='" + strDate + "' And 棟名称='"
                                        + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                            // ＤＢ読込み
                            DBTable = DBClass.Read2(connectionstring, DBSet2, strSQL);

                            // 検索件数が０以上なら、上書きモード
                            if (DBTable.Rows.Count > 0) {

                                // モードフラグ
                                sModeFlg = 1;

                                // データセットクリア
                                DBTable.Clear();
                            }
                            else
                                sModeFlg = 0;


                            // モードによって、発行するＳＱＬ文を変える
                            if (sModeFlg == 0) {
                                // INSERT(日報日付,棟名称,部屋名称,
                                // メータ値１日,２日,３日,４日,５日,６日,７日,８日,９日,１０日
                                // １１日,１２日,１３日,１４日,１５日,１６日,１７日,１８日,１９日,２０日
                                // ２１日,２２日,２３日,２４日,２５日,２６日,２７日,２８日,２９日,３０日
                                // ３１日);
                                strSQL = "INSERT INTO Ｄ月報データ VALUES('" + strDate + "','" + strToumei + "','" + strHeyamei + "',"
                                            + "0,0,0,0,0,0,0,0,0,0,"
                                            + "0,0,0,0,0,0,0,0,0,0,"
                                            + "0,0,0,0,0,0,0,0,0,0,"
                                            + "0);";

                                // 書込み
                                rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);

                                strTime = "メータ値" + Strings.StrConv(strDayBuf, VbStrConv.Wide) + "日";
                                strSQL = "UPDATE Ｄ月報データ SET " + strTime + "='" + Strings.Format(lngGeppouBuf, "#####0.#") + "' WHERE 月報日付='" + strDate + "' And 棟名称='" + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                                // 書込み
                                rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);
                            }
                            else {
                                // UPDATE
                                // 更新用文字列の作成
                                strTime = "メータ値" + Strings.StrConv(strDayBuf, VbStrConv.Wide) + "日";

                                strSQL = "UPDATE Ｄ月報データ SET " + strTime + "='" + Strings.Format(lngGeppouBuf, "#####0.#") + "' WHERE 月報日付='" + strDate + "' And 棟名称='" + strToumei + "' And 部屋名称='" + strHeyamei + "'";

                                // 書込み
                                rinf = DBClass.Write1(connectionstring, strSQL, sModeFlg);
                            }
                        }

                        // 日報ファイル作成
                        rinf = Make_DayReport(mINIData.mDayReportBasePath);

                        // バッファ更新
                        gDayBuf = Strings.Format(DateTime.Now, "%d");

                        #endregion

                    }

                    // 月替わり処理
                    if (gMonBuf != Strings.Format(DateTime.Now, "%M")) {
                        // 月報ファイル作成
                        // 月報強制印字機能追加により、関数の引数追加（０＝通常月報作成）
                        // rinf = Make_MonReport(mINIData.mMonReportBasePath)
                        rinf = Make_MonReport(mINIData.mMonReportBasePath, 0);

                        // バッファ更新
                        gMonBuf = Strings.Format(DateTime.Now, "%M");
                    }
                }

                #region Ｅメール版で自動送信モードの場合、締処理を行う

                if (mINIData.mEMODE == "1" & mINIData.mSENDMODE == "1") {
                    // 締日付比較用
                    strFinDateBuf = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH時");

                    if (mINIData.mNextFinDate == strFinDateBuf) {
                        // 締日作成処理
                        // 前回締日
                        mINIData.mLastFinDate = mINIData.mNowFinDate;

                        // 今回締日
                        mINIData.mNowFinDate = mINIData.mNextFinDate;

                        // 次回締日
                        mINIData.mNextFinDate = Format(DateAdd(DateInterval.Month, 1, (DateTime)mINIData.mNowFinDate), "yyyy年MM月dd日 HH時");
                        // 締日が月末の場合(DateAddで2月→3月になった場合、日付が28日or29日になっている為）
                        if (mINIData.mFinDay == "31") {
                            dtDate = (DateTime)mINIData.mNextFinDate;
                            strDay = CalenderCheck(dtDate);

                            strWorkBuf = Mid(mINIData.mNextFinDate, 1, 8);
                            strWorkBuf = strWorkBuf + strDay + Mid(mINIData.mNextFinDate, 11, 9);

                            dtDate = (DateTime)strWorkBuf;
                            mINIData.mNextFinDate = Strings.Format(dtDate, "yyyy年MM月dd日 HH時");
                        }

                        // ＩＮＩファイル更新
                        // 前回締日
                        rinf = WritePrivateProfileString(FINDATA, "LastFinDate", mINIData.mLastFinDate, INIFilePath);
                        // 今回締日
                        rinf = WritePrivateProfileString(FINDATA, "NowFinDate", mINIData.mNowFinDate, INIFilePath);
                        // 次回締日
                        rinf = WritePrivateProfileString(FINDATA, "NextFinDate", mINIData.mNextFinDate, INIFilePath);

                        // 2008-07-11 Add >>>
                        // 締日が変更された場合への対応として、締処理時に、次回送信日を再作成する。
                        // 実際には、今回締日＋指定日数となる。
                        // 次回送信日を更新
                        mINIData.mNextMailSend = DateAdd(DateInterval.Day, mINIData.mMailSendInterval, (DateTime)mINIData.mNowFinDate);

                        rinf = WritePrivateProfileString(FINDATA, "NextMailSend", mINIData.mNextMailSend, INIFilePath);
                        // <<<< Add End

                        // メール処理
                        rinf = Mail_Func(mINIData.mSENDMODE);

                        // ラベル更新
                        I_frmFeeCal.lblLastCal.Text = Format((DateTime)mINIData.mNowFinDate, "yyyy年MM月");

                        I_frmFeeCal.lblFeeDate.Text = Format((DateTime)mINIData.mLastFinDate, "yyyy年MM月dd日 HH時") + " ～ " + Format((DateTime)mINIData.mNowFinDate, "yyyy年MM月dd日 HH時");

                        I_frmFeeCal.lblNextFinDate.Text = Format((DateTime)mINIData.mNextFinDate, "yyyy年MM月dd日 HH時");
                    }

                    // 指定日付比較用
                    strFinDateBuf = Strings.Format(DateTime.Now, "yyyy年MM月dd日");
                    if ((DateTime)mINIData.mNextMailSend <= DateTime.Now) {
                        // 次回送信日を更新(次回締日+指定日数)
                        mINIData.mNextMailSend = DateAdd(DateInterval.Day, mINIData.mMailSendInterval, (DateTime)mINIData.mNextFinDate);

                        rinf = WritePrivateProfileString(FINDATA, "NextMailSend", mINIData.mNextMailSend, INIFilePath);

                        // ファイルパス作成
                        strPath = mINIData.mMailDataPath + mINIData.mMailFileName;

                        // メール送信(2:指定経過後送信)
                        rinf = Send_Mail(2, strPath);

                        // 2008-06-06 T.Isano Add >>>
                        // 送信結果判断処理
                        if (rinf == -1) {
                            gintRetFlg3 = 1;

                            // リトライカウンタ初期化
                            gintRetCnt = 0;

                            // ファイルパスをバッファへ退避（再送時に使用）
                            gstrMailFilePath = strPath;

                            // リトライ用タイマが起動していなければ、起動
                            if (tmRetry.Enabled == false)
                                tmRetry.Enabled = true;
                        }
                    }
                }

                #endregion


                // 年替わり処理
                if (gYearBuf != Strings.Format(DateTime.Now, "yyyy")) {
                    // ＤＢ削除
                    string strYearBuf;
                    // 日報データ削除
                    strYearBuf = Strings.Format(DateTime.DateAdd(DateInterval.Year, -3, DateTime.Now), "yyyy");
                    strSQL = "DELETE FROM Ｄ日報データ WHERE 日報日付 <= '" + strYearBuf + "年12月31日'";

                    // 書込み
                    rinf = DBClass.Write1(connectionstring, strSQL, 1);

                    // 月報データ削除
                    strYearBuf = Strings.Format(DateTime.DateAdd(DateInterval.Year, -6, DateTime.Now), "yyyy");
                    strSQL = "DELETE FROM Ｄ月報データ WHERE 月報日付 <= '" + strYearBuf + "年12月31日'";

                    // 書込み
                    rinf = DBClass.Write1(connectionstring, strSQL, 1);

                    // 課金月報データ削除
                    strYearBuf = Strings.Format(DateTime.DateAdd(DateInterval.Year, -6, DateTime.Now), "yyyy");
                    strSQL = "DELETE FROM Ｄ課金月報データ WHERE 課金月報日付 <= '" + strYearBuf + "年12月31日'";

                    // 書込み
                    rinf = DBClass.Write1(connectionstring, strSQL, 1);

                    gYearBuf = Strings.Format(DateTime.Now, "yyyy");
                }

                // 2007-04-04 >>>
                // 締日締時変更
                // 締日付比較用
                strFinDateBuf = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH時");

                if (mINIData.mNextFinDate == strFinDateBuf) {
                    // 前回締日
                    mINIData.mLastFinDate = mINIData.mNowFinDate;

                    // 今回締日
                    mINIData.mNowFinDate = mINIData.mNextFinDate;

                    // 次回締日
                    mINIData.mNextFinDate = Format(DateAdd(DateInterval.Month, 1, (DateTime)mINIData.mNowFinDate), "yyyy年MM月dd日 HH時");
                    // 締日が月末の場合(DateAddで2月→3月になった場合、日付が28日or29日になっている為）
                    if (mINIData.mFinDay == "31") {
                        dtDate = (DateTime)mINIData.mNextFinDate;
                        strDay = CalenderCheck(dtDate);

                        strWorkBuf = Mid(mINIData.mNextFinDate, 1, 8);
                        strWorkBuf = strWorkBuf + strDay + Mid(mINIData.mNextFinDate, 11, 9);

                        dtDate = (DateTime)strWorkBuf;
                        mINIData.mNextFinDate = Strings.Format(dtDate, "yyyy年MM月dd日 HH時");
                    }

                    // ＩＮＩファイル更新
                    // 前回締日
                    rinf = WritePrivateProfileString(FINDATA, "LastFinDate", mINIData.mLastFinDate, INIFilePath);
                    // 今回締日
                    rinf = WritePrivateProfileString(FINDATA, "NowFinDate", mINIData.mNowFinDate, INIFilePath);
                    // 次回締日
                    rinf = WritePrivateProfileString(FINDATA, "NextFinDate", mINIData.mNextFinDate, INIFilePath);

                    // ラベル更新
                    I_frmFeeCal.lblNextFinDate.Text = Format((DateTime)mINIData.mNextFinDate, "yyyy年MM月dd日 HH時");


                    // 収支演算データのお知らせを削除し、画面表示も消す
                    I_frmFeeCal.txtInform.Text = "";

                    // ＤＢ読込み
                    strSQL = "SELECT * FROM Ｄ収支演算データ WHERE 年月='" + Format((DateTime)mINIData.mLastFinDate, "yyyy年MM月dd日") + "'";

                    DBTable = DBClass.Read2(connectionstring, DBSet, strSQL);

                    // 検索件数が０以上なら処理
                    if (DBTable.Rows.Count > 0) {
                        row = DBTable.Rows(0);
                        // UPDATE

                        strSQL = "UPDATE Ｄ収支演算データ SET メッセージ='' WHERE 年月='" + Format((DateTime)mINIData.mLastFinDate, "yyyy年MM月dd日") + "'";

                        // 書込み
                        rinf = DBClass.Write1(connectionstring, strSQL, 1);

                        DBTable.Clear();
                    }
                }
            }
            // <<< 
            catch (Exception ex) {
                string strErr;

                // ログ出力文字列作成
                strErr = Strings.Format(DateTime.Now, "yyyy年MM月dd日 HH:mm:ss") + ":【frmMain】時刻タイマ処理でエラーが発生しました。" + "エラー内容：" + Constants.vbCrLf + ex.ToString() + Constants.vbCrLf;

                // ログ出力
                rinf = Log_Make(strErr);

                System.Environment.Exit(0);
            }
        }




    }
}
