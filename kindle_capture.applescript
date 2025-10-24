-- Kindle スクリーンショット自動撮影スクリプト
-- 設計書: plan.md v1.1
--
-- 機能: Kindleのページを自動的に連続でスクリーンショット撮影
-- 前提: Kindleは1つのみ起動、スペースキーでページ送り可能
-- 権限: アクセシビリティ権限＋スクリーン収録権限が必須

-- ============================================================
-- 設定可能な定数
-- ============================================================

property PAGE_TURN_DELAY : 1  -- ページめくり後の待機秒数（調整可）
property APP_NAME : "Kindle"
property SCREENSHOT_FOLDER_NAME : "kindle_screenshots"

-- ============================================================
-- メイン処理
-- ============================================================

try
    -- Step 1: Kindleウィンドウの自動検出
    set kindleWindowInfo to detectKindleWindow()

    -- Step 2: Kindleのアクティブ化＋最大化
    activateAndMaximizeKindle(kindleWindowInfo)

    -- Step 3: ユーザー操作待ち（最初のページ調整）
    waitForPageAdjustment()

    -- Step 4: 撮影設定入力（キャンセル・再入力対応）
    set pageCount to inputPageCount()

    -- Step 5: 保存先フォルダ作成
    set saveFolderPath to createScreenshotFolder()

    -- Step 6: スクリーンショット撮影ループ
    set captureResult to captureScreenshots(pageCount, saveFolderPath, kindleWindowInfo)
    set capturedCount to capturedCount of captureResult
    set earlyStopDetected to earlyStopDetected of captureResult

    -- Step 7: 完了通知
    showCompletionDialog(capturedCount, saveFolderPath, earlyStopDetected)

on error errMsg
    -- 予期しないエラー
    if errMsg does not contain "cancelled" then
        display dialog "エラーが発生しました：" & linefeed & errMsg buttons {"OK"} default button "OK" with icon caution
    end if
end try

-- ============================================================
-- 関数定義
-- ============================================================

-- Kindleウィンドウの自動検出
on detectKindleWindow()
    tell application "System Events"
        set kindleProcesses to (every process whose name contains APP_NAME)
    end tell

    -- Kindleプロセスの個数チェック
    if (count of kindleProcesses) = 0 then
        display dialog "Kindleアプリが見つかりません。Kindleを起動してください。" buttons {"OK"} default button "OK" with icon stop
        error "Kindle not found"
    end if

    if (count of kindleProcesses) > 1 then
        display dialog "Kindleウィンドウが複数開いています。1つだけ開いてください。" buttons {"OK"} default button "OK" with icon stop
        error "Multiple Kindle windows found"
    end if

    -- Kindleプロセス情報を取得
    set kindleProcess to item 1 of kindleProcesses
    set processName to name of kindleProcess

    tell application "System Events"
        tell process processName
            if (count of windows) = 0 then
                display dialog "Kindleで本が開かれていません。本を開いてください。" buttons {"OK"} default button "OK" with icon stop
                error "No Kindle window open"
            end if

            set mainWindow to window 1
            set windowName to name of mainWindow
        end tell
    end tell

    return {processName:processName, windowName:windowName}
end detectKindleWindow

-- Kindleのアクティブ化＋最大化
on activateAndMaximizeKindle(kindleWindowInfo)
    tell application APP_NAME
        activate
    end tell

    delay 0.5  -- アプリ起動待機

    -- ウィンドウを最大化
    tell application "System Events"
        set processName to processName of kindleWindowInfo
        tell process processName
            set frontmost to true

            if (count of windows) = 0 then error "No Kindle window available"

            set mainWindow to window 1

            -- 最小化されている場合は復元
            try
                if (value of attribute "AXMinimized" of mainWindow) is true then
                    set value of attribute "AXMinimized" of mainWindow to false
                    delay 0.3
                end if
            on error
                -- 属性が取得できない場合は無視
            end try

            try
                -- 画面サイズを設定
                set bounds of mainWindow to {0, 37, 1470, 919}
            end try
        end tell
    end tell

    delay 0.5  -- 描画完了待機
end activateAndMaximizeKindle

-- ユーザー操作待ち（最初のページ調整）
on waitForPageAdjustment()
    display dialog "Kindleの最初のページに合わせてOKを押してください" buttons {"OK"} default button "OK" with icon note
end waitForPageAdjustment

-- 撮影ページ数入力（キャンセル・再入力対応）
on inputPageCount()
    repeat
        try
            set dialResult to display dialog "何ページ撮影しますか？" default answer "10" buttons {"キャンセル", "OK"} default button "OK"
            set userInput to text returned of dialResult

            -- 空入力チェック
            if userInput = "" then
                display dialog "ページ数を入力してください。" buttons {"OK"} default button "OK" with icon caution
            else
                -- 数値検証
                try
                    set pageNum to userInput as integer

                    if pageNum < 1 then
                        display dialog "1以上の数値を入力してください。" buttons {"OK"} default button "OK" with icon caution
                    else
                        return pageNum
                    end if
                on error
                    display dialog "1以上の整数を入力してください。" buttons {"OK"} default button "OK" with icon caution
                end try
            end if
        on error errMsg
            -- Cancel ボタンが押された場合
            if button returned of dialResult = "キャンセル" then
                error "User cancelled the operation"
            end if
        end try
    end repeat
end inputPageCount

-- 保存先フォルダ作成
on createScreenshotFolder()
    try
        -- デスクトップパスを動的に取得（iCloud Desktop対応）
        set desktopPath to POSIX path of (path to desktop folder)

        -- タイムスタンプ生成（YYYYMMDD_HHMMSS形式）
        set timestamp to do shell script "date +%Y%m%d_%H%M%S"

        -- フォルダパス構築
        set basePath to desktopPath & SCREENSHOT_FOLDER_NAME & "/"
        set folderPath to basePath & timestamp & "/"

        -- フォルダ作成
        do shell script "mkdir -p " & quoted form of folderPath

        return folderPath
    on error errMsg
        display dialog "スクリーンショット保存フォルダを作成できませんでした。" & linefeed & "デスクトップのアクセス権限を確認してください。" & linefeed & linefeed & errMsg buttons {"OK"} default button "OK" with icon stop
        error "Failed to create screenshot folder"
    end try
end createScreenshotFolder

-- スクリーンショット撮影ループ
on captureScreenshots(pageCount, folderPath, kindleWindowInfo)
    set processName to processName of kindleWindowInfo
    set windowBounds to getKindleWindowBounds(processName)

    -- ウィンドウ座標を設定（screencapture -R用）
    set {x1, y1, x2, y2} to windowBounds
    set boundsStr to (x1 as string) & "," & (y1 as string) & "," & ((x2 - x1) as string) & "," & ((y2 - y1) as string)

    set previousHash to missing value
    set duplicateStreak to 0
    set capturedCount to 0
    set earlyStopDetected to false

    repeat with i from 1 to pageCount
        try
            -- ゼロパディング（001, 002, ...）
            set paddedNumber to text -3 thru -1 of ("000" & i)
            set fileName to "page_" & paddedNumber & ".png"
            set filePath to folderPath & fileName

            -- スクリーンショット撮影（-R で座標範囲指定）
            do shell script "screencapture -R " & boundsStr & " -x " & quoted form of filePath

            -- ハッシュで同一ページ判定
            set currentHash to do shell script "shasum -a 256 " & quoted form of filePath & " | awk '{print $1}'"
            if previousHash is missing value then
                set duplicateStreak to 1
            else if currentHash = previousHash then
                set duplicateStreak to duplicateStreak + 1
            else
                set duplicateStreak to 1
            end if

            if duplicateStreak ≥ 2 then
                set earlyStopDetected to true
                try
                    -- 2回目の重複キャプチャは残さない
                    do shell script "rm -f " & quoted form of filePath
                    log "Duplicate page detected. Removed " & fileName
                end try
                exit repeat
            else
                set capturedCount to capturedCount + 1
                set previousHash to currentHash

                -- 最後のページでない場合のみページめくり
                if i < pageCount then
                    -- スペースキーでページめくり
                    tell application "System Events"
                        tell process processName
                            keystroke space
                        end tell
                    end tell

                    -- ページ描画待機（PAGE_TURN_DELAY秒）
                    delay PAGE_TURN_DELAY
                end if
            end if

        on error errMsg
            -- エラーが発生してもスクリプト継続
            log "Screenshot " & fileName & " failed: " & errMsg
        end try
    end repeat

    return {capturedCount:capturedCount, earlyStopDetected:earlyStopDetected}
end captureScreenshots

-- 現在のKindleウィンドウ座標を取得
on getKindleWindowBounds(processName)
    tell application "System Events"
        tell process processName
            set frontmost to true
            delay 0.2

            if (count of windows) = 0 then error "No Kindle window available"

            set mainWindow to window 1

            -- 最小化されている場合は復元
            try
                if (value of attribute "AXMinimized" of mainWindow) is true then
                    set value of attribute "AXMinimized" of mainWindow to false
                    delay 0.3
                end if
            on error
                -- 属性が取得できない場合は無視
            end try

            -- bounds取得は失敗することがあるのでリトライを用意
            set windowBounds to missing value
            repeat with attempt from 1 to 3
                try
                    set windowBounds to bounds of mainWindow
                    exit repeat
                on error
                    try
                        set {xPos, yPos} to position of mainWindow
                        set {wSize, hSize} to size of mainWindow
                        set xLeft to round xPos rounding down
                        set yTop to round yPos rounding down
                        set widthPixels to round wSize rounding down
                        set heightPixels to round hSize rounding down
                        set windowBounds to {xLeft, yTop, xLeft + widthPixels, yTop + heightPixels}
                        exit repeat
                    on error
                        if attempt = 3 then error "Failed to obtain Kindle window bounds"
                        delay 0.3
                    end try
                end try
            end repeat
        end tell
    end tell

    return windowBounds
end getKindleWindowBounds

on showCompletionDialog(capturedCount, folderPath, earlyStopDetected)
    set capturedCountText to capturedCount as string
    set message to capturedCountText & "枚のスクリーンショットを保存しました。" & linefeed & linefeed & "保存先: " & folderPath
    if earlyStopDetected is true then
        set message to message & linefeed & "(2回連続で同じページが検出されたため、撮影を終了しました)"
    end if

    display dialog message buttons {"フォルダを開く", "OK"} default button "OK" with icon note

    -- 「フォルダを開く」ボタンが押された場合
    if button returned of result = "フォルダを開く" then
        tell application "Finder"
            activate
            open (folderPath as POSIX file)
        end tell
    end if
end showCompletionDialog
