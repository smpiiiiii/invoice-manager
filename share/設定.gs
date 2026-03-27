/**
 * 📋 請求書自動管理システム — 設定ファイル
 *
 * ★ 初回セットアップ:
 *   1. GEMINI_API_KEY を設定（下の手順参照）
 *   2. processInvoices() を手動実行して権限を承認
 *   3. 自動でスプレッドシート・Driveフォルダが作成されます
 *
 * ★ API Key取得方法:
 *   https://aistudio.google.com/apikey にアクセス
 *   → 「APIキーを作成」→ キーをコピー → 下に貼り付け
 */

// ===== 必須設定 =====

/** Gemini API Key（上記URLで取得してここに貼り付け） */
var GEMINI_API_KEY = 'ここにAPIキーを貼り付け';

/** Gemini API URL（変更不要） */
var GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/';

// ===== 自動生成される設定（初回実行時に自動セット、手動変更不要） =====

/** スプレッドシートID（空なら自動作成） */
var SPREADSHEET_ID = '';

/** Google Drive 請求書保存フォルダID（空なら自動作成） */
var DRIVE_FOLDER_ID = '';

// ===== カスタマイズ可能な設定 =====

/** Gmail検索クエリ（PDF添付メール全検索） */
var GMAIL_QUERY = 'has:attachment filename:pdf';

/** 処理済みラベル名（二重処理防止用） */
var PROCESSED_LABEL = '請求書処理済';

/** 処理対象開始日（これ以降のメールを処理） */
var START_DATE = '2026/02/01';

/** 1回の実行で処理するメール数の上限（GAS実行時間制限対策） */
var BATCH_SIZE = 30;

/** Gemini APIモデル（変更不要） */
var GEMINI_MODEL = 'gemini-2.5-flash';

/** スプレッドシート名 */
var SHEET_NAME = '請求書管理';

/** Driveフォルダ名 */
var FOLDER_NAME = '請求書';
