// スタイルが当たっていないテキストを見つけて、近いスタイルを適用するプラグイン

// 近いスタイルがあるテキストノードのIDを保存
let hasNearestTextNodeIds = [];
// 近いスタイルがないテキストノードのIDを保存
let noNearestTextNodeIds = [];

// UIが表示されているかどうかのフラグ
let isUIVisible = false;

// UIを開く関数
function openUI() {
    if (!isUIVisible) {
        figma.showUI(__html__, { width: 320, height: 900 });
        isUIVisible = true;
        
        // 選択状態の変更を監視
        figma.on('selectionchange', async () => {
            await scanStyles();
        });
        
        // 初期状態をチェック
        scanStyles();
    }
}

// プラグイン起動時のコマンドをチェック
const command = figma.command;

if (command === 'open-panel' || command === undefined || command === '') {
    // パネルを開くコマンド、またはコマンドがない場合（直接実行）はUIを開く
    openUI();
} else if (command === 'apply-larger-style') {
    // 文字サイズの大きいスタイルを適用
    applyLargerStyle().then(() => {
        figma.closePlugin();
    });
} else if (command === 'apply-smaller-style') {
    // 文字サイズの小さいスタイルを適用
    applySmallerStyle().then(() => {
        figma.closePlugin();
    });
}

figma.ui.onmessage = async (msg) => {
    if (msg.type === 'scan-styles') {
        await scanStyles();
    } else if (msg.type === 'select-has-nearest') {
        await selectHasNearestNodes();
    } else if (msg.type === 'select-no-nearest') {
        await selectNoNearestNodes();
    } else if (msg.type === 'apply-nearest-styles') {
        await applyNearestStyles();
    } else if (msg.type === 'apply-closest-styles') {
        await applyClosestStyles();
    } else if (msg.type === 'create-new-styles') {
        await createNewStyles();
    } else if (msg.type === 'check-unused-styles') {
        await checkUnusedStyles();
    } else if (msg.type === 'delete-unused-styles') {
        await deleteUnusedStyles();
    } else if (msg.type === 'set-line-height-150') {
        await setLineHeightTo150();
    } else if (msg.type === 'get-styles-list') {
        await getStylesList();
    } else if (msg.type === 'bulk-edit-styles') {
        await bulkEditStyles(msg.property, msg.value, msg.styleIds);
    } else if (msg.type === 'select-nodes-with-styles') {
        await selectNodesWithStyles(msg.styleIds);
    } else if (msg.type === 'cancel') {
        figma.closePlugin();
    }
};

// 選択状態をチェック
function checkSelection() {
    const selection = figma.currentPage.selection;
    return selection.length > 0;
}

// スキャン機能：スタイルの状態を分析する（適用はしない）
async function scanStyles() {
    // UIが開いていない場合は何もしない
    if (!isUIVisible) {
        return;
    }
    
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
        figma.ui.postMessage({
            type: 'scan-result',
            hasSelection: false
        });
        return;
    }

    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }

    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();

    let appliedCount = 0; // スタイル適用済み
    let notAppliedCount = 0; // スタイル未適用
    
    // 近いスタイルがあるテキストノードのIDをリセット
    hasNearestTextNodeIds = [];
    // 近いスタイルがないテキストノードのIDをリセット
    noNearestTextNodeIds = [];
    
    // プロパティの種類をカウントするためのSet（キーはプロパティを文字列化したもの）
    const hasNearestPropsSet = new Set();
    const noNearestPropsSet = new Set();

    for (const textNode of textNodes) {
        // スタイルが既に当たっているかチェック
        if (hasTextStyle(textNode)) {
            appliedCount++;
            continue;
        }

        // スタイルが当たっていない
        notAppliedCount++;

        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);

        if (!textProps) {
            // プロパティを取得できないテキストノードのIDを保存
            noNearestTextNodeIds.push(textNode.id);
            noNearestPropsSet.add('unknown');
            continue;
        }

        // プロパティをキー文字列に変換
        const propsKey = createPropsKey(textProps);

        // 近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);

        if (nearestStyle) {
            hasNearestPropsSet.add(propsKey);
            // 近いスタイルがあるテキストノードのIDを保存
            hasNearestTextNodeIds.push(textNode.id);
        } else {
            noNearestPropsSet.add(propsKey);
            // 近いスタイルがないテキストノードのIDを保存
            noNearestTextNodeIds.push(textNode.id);
        }
    }
    
    // プロパティの種類数をカウント
    const hasNearestCount = hasNearestPropsSet.size;
    const noNearestCount = noNearestPropsSet.size;

    // 結果を送信
    figma.ui.postMessage({
        type: 'scan-result',
        hasSelection: true,
        applied: appliedCount,
        notApplied: notAppliedCount,
        hasNearest: hasNearestCount,
        noNearest: noNearestCount
    });
}

// 近いスタイルがあるテキストノードを選択
async function selectHasNearestNodes() {
    // 現在選択されているノードから開始（またはページ全体）
    const currentSelection = figma.currentPage.selection;
    let rootNodes = [];
    
    if (currentSelection.length > 0) {
        // 現在選択されているノードをルートとして使用
        rootNodes = currentSelection;
    } else {
        // 選択がない場合は、ページ全体のフレームを対象にする
        // ただし、これは大量のノードになる可能性があるので、選択がある場合のみ動作させる
        return;
    }
    
    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of rootNodes) {
        collectTextNodes(node, textNodes);
    }
    
    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    // 近いスタイルがあるテキストノードを収集
    const nodesToSelect = [];
    
    for (const textNode of textNodes) {
        // スタイルが既に当たっている場合はスキップ
        if (hasTextStyle(textNode)) {
            continue;
        }
        
        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);
        
        if (!textProps) {
            continue;
        }
        
        // 近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);
        
        if (nearestStyle) {
            // 近いスタイルがあるテキストノードを選択対象に追加
            nodesToSelect.push(textNode);
        }
    }
    
    // ノードを選択
    if (nodesToSelect.length > 0) {
        try {
            figma.currentPage.selection = nodesToSelect;
        } catch (error) {
            console.error('Error selecting nodes:', error);
        }
    }
}

// 近いスタイルがないテキストノードを選択
async function selectNoNearestNodes() {
    // 現在選択されているノードから開始（またはページ全体）
    const currentSelection = figma.currentPage.selection;
    let rootNodes = [];
    
    if (currentSelection.length > 0) {
        // 現在選択されているノードをルートとして使用
        rootNodes = currentSelection;
    } else {
        // 選択がない場合は、ページ全体のフレームを対象にする
        // ただし、これは大量のノードになる可能性があるので、選択がある場合のみ動作させる
        return;
    }
    
    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of rootNodes) {
        collectTextNodes(node, textNodes);
    }
    
    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    // 近いスタイルがないテキストノードを収集
    const nodesToSelect = [];
    
    for (const textNode of textNodes) {
        // スタイルが既に当たっている場合はスキップ
        if (hasTextStyle(textNode)) {
            continue;
        }
        
        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);
        
        if (!textProps) {
            continue;
        }
        
        // 近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);
        
        if (!nearestStyle) {
            // 近いスタイルがないテキストノードを選択対象に追加
            nodesToSelect.push(textNode);
        }
    }
    
    // ノードを選択
    if (nodesToSelect.length > 0) {
        try {
            figma.currentPage.selection = nodesToSelect;
        } catch (error) {
            console.error('Error selecting nodes:', error);
        }
    }
}

// 近いスタイルを適用する（近いスタイルがあるものだけ）
async function applyNearestStyles() {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
        return;
    }

    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }

    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();

    let appliedCount = 0;
    let skippedCount = 0;

    for (const textNode of textNodes) {
        // スタイルが既に当たっているかチェック
        if (hasTextStyle(textNode)) {
            continue; // スタイルが当たっている場合はスキップ
        }

        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);

        if (!textProps) {
            skippedCount++;
            continue;
        }

        // 近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);

        if (nearestStyle) {
            // スタイルを適用
            try {
                await textNode.setTextStyleIdAsync(nearestStyle.id);
                appliedCount++;
            } catch (error) {
                console.error('Error applying text style:', error);
                skippedCount++;
            }
        } else {
            // 近いスタイルがない場合はスキップ
            skippedCount++;
        }
    }

    // 通知を表示
    if (appliedCount > 0) {
        figma.ui.postMessage({
            type: 'notification',
            message: `${appliedCount}件 近いスタイルを適用しました`,
            color: 'green'
        });
    }
    
    // スキャンを再実行して結果を更新
    setTimeout(() => {
        scanStyles();
    }, 500);
}

// スタイルを新規作成する（スタイル未適用の全テキストに対して）
async function createNewStyles() {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
        return;
    }

    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }

    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();

    let createdCount = 0;
    let skippedCount = 0;

    for (const textNode of textNodes) {
        // スタイルが既に当たっているかチェック
        if (hasTextStyle(textNode)) {
            continue; // スタイルが当たっている場合はスキップ
        }

        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);

        if (!textProps) {
            skippedCount++;
            continue;
        }

        // 新規作成（近いスタイルの有無に関係なく）
        try {
            const newStyle = await createTextStyleFromNode(textNode, textProps);
            if (newStyle) {
                textStyles.push(newStyle); // 作成したスタイルをリストに追加
                await textNode.setTextStyleIdAsync(newStyle.id);
                createdCount++;
            }
        } catch (error) {
            console.error('Error creating text style:', error);
            skippedCount++;
        }
    }

    // 通知を表示
    if (createdCount > 0) {
        figma.ui.postMessage({
            type: 'notification',
            message: `${createdCount}件 新しいスタイルを作成しました`,
            color: 'green'
        });
    }
    
    // スキャンを再実行して結果を更新
    setTimeout(() => {
        scanStyles();
    }, 500);
}

// 一番近いスタイルを適用する（条件を無視して、近いスタイルがないものに対して適用）
async function applyClosestStyles() {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
        return;
    }

    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }

    // ドキュメント内のすべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();

    let appliedCount = 0;
    let skippedCount = 0;
    let differentFamilyCount = 0;

    for (const textNode of textNodes) {
        // スタイルが既に当たっているかチェック
        if (hasTextStyle(textNode)) {
            continue; // スタイルが当たっている場合はスキップ
        }

        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);

        if (!textProps) {
            skippedCount++;
            continue;
        }

        // 通常の条件で近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);

        // 通常の条件で近いスタイルがない場合のみ、一番近いスタイルを探す
        if (!nearestStyle) {
            const closestResult = await findClosestStyle(textProps, textStyles);
            
            if (closestResult) {
                // 一番近いスタイルを適用
                try {
                    await textNode.setTextStyleIdAsync(closestResult.style.id);
                    appliedCount++;
                    
                    // フォントファミリーが異なる場合はカウント
                    if (closestResult.isDifferentFamily) {
                        differentFamilyCount++;
                    }
                } catch (error) {
                    console.error('Error applying closest text style:', error);
                    skippedCount++;
                }
            } else {
                skippedCount++;
            }
        } else {
            // 通常の条件で近いスタイルがある場合はスキップ
            skippedCount++;
        }
    }

    // 通知を表示
    if (appliedCount > 0) {
        let message = `${appliedCount}件 まだ近いスタイルを適用しました`;
        
        // フォントファミリーが異なるスタイルを適用した場合は警告
        if (differentFamilyCount > 0) {
            message += `<br>⚠️ ${differentFamilyCount}件はフォントファミリーが異なるスタイルを適用しました`;
        }
        
        figma.ui.postMessage({
            type: 'notification',
            message: message,
            color: 'green'
        });
    }
    
    // スキャンを再実行して結果を更新
    setTimeout(() => {
        scanStyles();
    }, 500);
}

// テキストノードを再帰的に収集
function collectTextNodes(node, textNodes) {
    if (node.type === 'TEXT') {
        textNodes.push(node);
    } else if ('children' in node) {
        for (const child of node.children) {
            collectTextNodes(child, textNodes);
        }
    }
}

// テキストにスタイルが当たっているかチェック
function hasTextStyle(textNode) {
    // textStyleIdがシンボル（mixed）の場合は一部にスタイルが当たっている
    // 空文字列またはnullの場合はスタイルが当たっていない
    const styleId = textNode.textStyleId;

    if (styleId === figma.mixed) {
        // 一部にスタイルが当たっている場合も対象に含める
        return false;
    }

    if (styleId === '' || styleId === null) {
        return false;
    }

    return true;
}

// プロパティをキー文字列に変換（重複判定用）
function createPropsKey(textProps) {
    // 行の高さを正規化
    let lineHeightValue = 'auto';
    if (textProps.lineHeight && textProps.lineHeight.unit !== 'AUTO') {
        if (textProps.lineHeight.unit === 'PIXELS') {
            lineHeightValue = `${textProps.lineHeight.value}px`;
        } else if (textProps.lineHeight.unit === 'PERCENT') {
            lineHeightValue = `${textProps.lineHeight.value}%`;
        }
    }
    
    return `${textProps.fontFamily}|${textProps.fontWeight}|${textProps.fontSize}|${lineHeightValue}`;
}

// テキストノードのプロパティを取得（代表値）
function getTextProperties(textNode) {
    try {
        let fontFamily, fontWeight, fontSize, lineHeight;

        // フォント名の取得
        const fontName = textNode.fontName;
        if (fontName === figma.mixed) {
            // mixedの場合は最初の文字のプロパティを取得
            fontFamily = textNode.getRangeFontName(0, 1).family;
            fontWeight = textNode.getRangeFontName(0, 1).style;
        } else {
            fontFamily = fontName.family;
            fontWeight = fontName.style;
        }

        // フォントサイズの取得
        if (textNode.fontSize === figma.mixed) {
            fontSize = textNode.getRangeFontSize(0, 1);
        } else {
            fontSize = textNode.fontSize;
        }

        // 行の高さの取得
        if (textNode.lineHeight === figma.mixed) {
            lineHeight = textNode.getRangeLineHeight(0, 1);
        } else {
            lineHeight = textNode.lineHeight;
        }

        return {
            fontFamily,
            fontWeight,
            fontSize,
            lineHeight
        };
    } catch (error) {
        console.error('Error getting text properties:', error);
        return null;
    }
}

// 近いスタイルを見つける
async function findNearestStyle(textProps, textStyles) {
    let nearestStyle = null;

    for (const style of textStyles) {
        // スタイルのプロパティを取得
        const styleFontName = style.fontName;
        const styleFontSize = style.fontSize;
        const styleLineHeight = style.lineHeight;

        // フォントファミリーが同じかチェック
        if (styleFontName.family !== textProps.fontFamily) {
            continue;
        }

        // フォントウェイト（style）が1段階以内かチェック
        const textWeightNum = fontWeightToNumber(textProps.fontWeight);
        const styleWeightNum = fontWeightToNumber(styleFontName.style);
        const weightDiff = Math.abs(styleWeightNum - textWeightNum);
        if (weightDiff > 100) {
            continue;
        }

        // フォントサイズが±1以内かチェック
        const sizeDiff = Math.abs(styleFontSize - textProps.fontSize);
        if (sizeDiff > 1) {
            continue;
        }

        // 行の高さが20%以内かチェック
        // AUTO（デフォルト）の場合とピクセル値の場合で処理を分ける
        if (textProps.lineHeight && textProps.lineHeight.unit !== 'AUTO' &&
            styleLineHeight && styleLineHeight.unit !== 'AUTO') {

            let textLineHeightValue, styleLineHeightValue;

            // lineHeightの値を取得（unitによって処理を分ける）
            if (textProps.lineHeight.unit === 'PIXELS') {
                textLineHeightValue = textProps.lineHeight.value;
            } else if (textProps.lineHeight.unit === 'PERCENT') {
                textLineHeightValue = textProps.fontSize * (textProps.lineHeight.value / 100);
            } else {
                textLineHeightValue = textProps.fontSize * 1.2; // デフォルト値
            }

            if (styleLineHeight.unit === 'PIXELS') {
                styleLineHeightValue = styleLineHeight.value;
            } else if (styleLineHeight.unit === 'PERCENT') {
                styleLineHeightValue = styleFontSize * (styleLineHeight.value / 100);
            } else {
                styleLineHeightValue = styleFontSize * 1.2; // デフォルト値
            }

            // 20%以上差があるかチェック
            const heightDiffPercent = Math.abs(styleLineHeightValue - textLineHeightValue) / textLineHeightValue;
            if (heightDiffPercent > 0.2) {
                continue;
            }
        }

        // 条件に合致するスタイルが見つかった
        // より近いサイズのスタイルを優先
        if (!nearestStyle) {
            nearestStyle = {
                style,
                sizeDiff
            };
        } else {
            if (sizeDiff < nearestStyle.sizeDiff) {
                nearestStyle = {
                    style,
                    sizeDiff
                };
            }
        }
    }

    return nearestStyle ? nearestStyle.style : null;
}

// フォントウェイト名を数値に変換
function fontWeightToNumber(weight) {
    const weightMap = {
        'Thin': 100,
        'ExtraLight': 200,
        'UltraLight': 200,
        'Light': 300,
        'Regular': 400,
        'Normal': 400,
        'Medium': 500,
        'SemiBold': 600,
        'DemiBold': 600,
        'Bold': 700,
        'ExtraBold': 800,
        'UltraBold': 800,
        'Black': 900,
        'Heavy': 900
    };
    
    // 大文字小文字を区別せずに検索
    const normalizedWeight = weight.charAt(0).toUpperCase() + weight.slice(1);
    
    // 完全一致を探す
    if (weightMap[normalizedWeight] !== undefined) {
        return weightMap[normalizedWeight];
    }
    
    // 部分一致を探す（例: "SemiBold" が "Semi Bold" として来る場合）
    for (const key in weightMap) {
        if (normalizedWeight.includes(key) || key.includes(normalizedWeight)) {
            return weightMap[key];
        }
    }
    
    // デフォルト値（見つからない場合はRegularとして扱う）
    return 400;
}

// 条件を無視して一番近いスタイルを見つける（サイズとウェイトが最も近いもの）
// フォントファミリーが一致するものを優先し、なければ異なるファミリーからも探す
async function findClosestStyle(textProps, textStyles) {
    let closestStyleSameFamily = null;
    let minTotalDiffSameFamily = Infinity;
    let closestStyleAnyFamily = null;
    let minTotalDiffAnyFamily = Infinity;

    // テキストのウェイトを数値に変換
    const textWeightNum = fontWeightToNumber(textProps.fontWeight);

    for (const style of textStyles) {
        // スタイルのプロパティを取得
        const styleFontName = style.fontName;
        const styleFontSize = style.fontSize;

        // フォントサイズの差を計算
        const sizeDiff = Math.abs(styleFontSize - textProps.fontSize);
        
        // フォントウェイトの差を計算
        const styleWeightNum = fontWeightToNumber(styleFontName.style);
        const weightDiff = Math.abs(styleWeightNum - textWeightNum);
        
        // サイズとウェイトの合計差を計算（サイズの差を優先するため、ウェイトの差は小さめの係数をかける）
        const totalDiff = sizeDiff + (weightDiff / 100) * 0.5;

        // フォントファミリーが同じ場合
        if (styleFontName.family === textProps.fontFamily) {
            if (totalDiff < minTotalDiffSameFamily) {
                minTotalDiffSameFamily = totalDiff;
                closestStyleSameFamily = style;
            }
        }
        
        // 全てのファミリーから探す（フォールバック用）
        if (totalDiff < minTotalDiffAnyFamily) {
            minTotalDiffAnyFamily = totalDiff;
            closestStyleAnyFamily = style;
        }
    }

    // 同じファミリーのスタイルがあればそれを返す、なければ異なるファミリーのスタイルを返す
    if (closestStyleSameFamily) {
        return { style: closestStyleSameFamily, isDifferentFamily: false };
    } else if (closestStyleAnyFamily) {
        return { style: closestStyleAnyFamily, isDifferentFamily: true };
    }
    
    return null;
}

// テキストノードから新規スタイルを作成
async function createTextStyleFromNode(textNode, textProps) {
    // スタイル名を生成（フォントファミリー/ウェイト/サイズ）
    const styleName = `${textProps.fontFamily}/${textProps.fontWeight}/${textProps.fontSize}`;

    // 新しいテキストスタイルを作成
    const newStyle = figma.createTextStyle();
    newStyle.name = styleName;

    // フォントを読み込む
    await figma.loadFontAsync({
        family: textProps.fontFamily,
        style: textProps.fontWeight
    });

    // スタイルのプロパティを設定
    newStyle.fontName = {
        family: textProps.fontFamily,
        style: textProps.fontWeight
    };
    newStyle.fontSize = textProps.fontSize;

    // テキストノードから他のプロパティもコピー
    try {
        // 文字間隔
        if (textNode.letterSpacing !== figma.mixed && textNode.letterSpacing !== undefined) {
            try {
                newStyle.letterSpacing = textNode.letterSpacing;
            } catch (e) {
                console.log('Could not set letterSpacing');
            }
        }

        // 行の高さ
        if (textNode.lineHeight !== figma.mixed && textNode.lineHeight !== undefined) {
            try {
                newStyle.lineHeight = textNode.lineHeight;
            } catch (e) {
                console.log('Could not set lineHeight');
            }
        }

        // テキストの整列
        if (textNode.textAlignHorizontal !== figma.mixed && textNode.textAlignHorizontal !== undefined) {
            try {
                newStyle.textAlignHorizontal = textNode.textAlignHorizontal;
            } catch (e) {
                console.log('Could not set textAlignHorizontal');
            }
        }

        if (textNode.textAlignVertical !== figma.mixed && textNode.textAlignVertical !== undefined) {
            try {
                newStyle.textAlignVertical = textNode.textAlignVertical;
            } catch (e) {
                console.log('Could not set textAlignVertical');
            }
        }

        // テキストの大文字・小文字の設定
        if (textNode.textCase !== figma.mixed && textNode.textCase !== undefined) {
            try {
                newStyle.textCase = textNode.textCase;
            } catch (e) {
                console.log('Could not set textCase');
            }
        }

        // テキストの装飾
        if (textNode.textDecoration !== figma.mixed && textNode.textDecoration !== undefined) {
            try {
                newStyle.textDecoration = textNode.textDecoration;
            } catch (e) {
                console.log('Could not set textDecoration');
            }
        }

        // 段落の間隔
        if (textNode.paragraphSpacing !== figma.mixed && textNode.paragraphSpacing !== undefined) {
            try {
                newStyle.paragraphSpacing = textNode.paragraphSpacing;
            } catch (e) {
                console.log('Could not set paragraphSpacing');
            }
        }

        // 段落のインデント
        if (textNode.paragraphIndent !== figma.mixed && textNode.paragraphIndent !== undefined) {
            try {
                newStyle.paragraphIndent = textNode.paragraphIndent;
            } catch (e) {
                console.log('Could not set paragraphIndent');
            }
        }
    } catch (error) {
        console.error('Error setting additional style properties:', error);
    }

    return newStyle;
}

// 未使用スタイルをチェック
async function checkUnusedStyles() {
    // すべてのローカルテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();

    if (textStyles.length === 0) {
        figma.ui.postMessage({
            type: 'unused-styles-result',
            totalStyles: 0,
            unusedStyles: []
        });
        return;
    }

    // 使用されているスタイルIDのセット
    const usedStyleIds = new Set();

    // ページ全体のテキストノードを収集
    const allTextNodes = [];
    collectTextNodes(figma.currentPage, allTextNodes);

    // 各テキストノードで使用されているスタイルIDを記録
    for (const textNode of allTextNodes) {
        const styleId = textNode.textStyleId;

        if (styleId && styleId !== figma.mixed && styleId !== '' && styleId !== null) {
            usedStyleIds.add(styleId);
        }

        // mixedの場合は各範囲のスタイルをチェック
        if (styleId === figma.mixed) {
            try {
                const length = textNode.characters.length;
                for (let i = 0; i < length; i++) {
                    const rangeStyleId = textNode.getRangeTextStyleId(i, i + 1);
                    if (rangeStyleId && rangeStyleId !== '' && rangeStyleId !== figma.mixed) {
                        usedStyleIds.add(rangeStyleId);
                    }
                }
            } catch (error) {
                console.error('Error checking mixed styles:', error);
            }
        }
    }

    // 未使用のスタイルを見つける
    const unusedStyles = [];
    for (const style of textStyles) {
        if (!usedStyleIds.has(style.id)) {
            unusedStyles.push({
                name: style.name,
                id: style.id
            });
        }
    }

    // 結果を送信
    figma.ui.postMessage({
        type: 'unused-styles-result',
        totalStyles: textStyles.length,
        usedStyles: usedStyleIds.size,
        unusedStyles: unusedStyles,
        totalTextNodes: allTextNodes.length
    });
}

// 未使用スタイルを削除
async function deleteUnusedStyles() {
    // すべてのローカルテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    if (textStyles.length === 0) {
        figma.ui.postMessage({
            type: 'delete-unused-styles-result',
            deletedCount: 0,
            error: null
        });
        return;
    }
    
    // 使用されているスタイルIDのセット
    const usedStyleIds = new Set();
    
    // ページ全体のテキストノードを収集
    const allTextNodes = [];
    collectTextNodes(figma.currentPage, allTextNodes);
    
    // 各テキストノードで使用されているスタイルIDを記録
    for (const textNode of allTextNodes) {
        const styleId = textNode.textStyleId;
        
        if (styleId && styleId !== figma.mixed && styleId !== '' && styleId !== null) {
            usedStyleIds.add(styleId);
        }
        
        // mixedの場合は各範囲のスタイルをチェック
        if (styleId === figma.mixed) {
            try {
                const length = textNode.characters.length;
                for (let i = 0; i < length; i++) {
                    const rangeStyleId = textNode.getRangeTextStyleId(i, i + 1);
                    if (rangeStyleId && rangeStyleId !== '' && rangeStyleId !== figma.mixed) {
                        usedStyleIds.add(rangeStyleId);
                    }
                }
            } catch (error) {
                console.error('Error checking mixed styles:', error);
            }
        }
    }
    
    // 未使用のスタイルを削除
    let deletedCount = 0;
    const errors = [];
    
    for (const style of textStyles) {
        if (!usedStyleIds.has(style.id)) {
            try {
                style.remove();
                deletedCount++;
            } catch (error) {
                console.error(`Error deleting style ${style.name}:`, error);
                errors.push({
                    name: style.name,
                    error: error.message || 'Unknown error'
                });
            }
        }
    }
    
    // 結果を送信
    figma.ui.postMessage({
        type: 'delete-unused-styles-result',
        deletedCount: deletedCount,
        errors: errors,
        error: errors.length > 0 ? `${errors.length}件のスタイルで削除エラーが発生しました` : null
    });
}

// 行間がAutoになっているスタイルを全て150%に変更
async function setLineHeightTo150() {
    // すべてのローカルテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    if (textStyles.length === 0) {
        figma.ui.postMessage({
            type: 'set-line-height-result',
            updatedCount: 0,
            error: null
        });
        return;
    }
    
    let updatedCount = 0;
    const errors = [];
    
    for (const style of textStyles) {
        try {
            // 行間がAUTOかチェック
            const lineHeight = style.lineHeight;
            
            // lineHeightがundefined、null、またはunitが'AUTO'の場合
            if (!lineHeight || (lineHeight.unit === 'AUTO')) {
                // フォントを読み込む（行間を設定する前に必要）
                try {
                    const fontName = style.fontName;
                    await figma.loadFontAsync({
                        family: fontName.family,
                        style: fontName.style
                    });
                    
                    // 行間を150%に設定
                    style.lineHeight = {
                        unit: 'PERCENT',
                        value: 150
                    };
                    updatedCount++;
                } catch (setError) {
                    // 設定に失敗した場合はエラーを記録
                    console.error(`Error setting line height for style ${style.name}:`, setError);
                    const errorMessage = setError.message || setError.toString() || 'Failed to set line height';
                    errors.push({
                        name: style.name,
                        error: errorMessage
                    });
                }
            }
        } catch (error) {
            console.error(`Error updating line height for style ${style.name}:`, error);
            errors.push({
                name: style.name,
                error: error.message || 'Unknown error'
            });
        }
    }
    
    // 結果を送信
    figma.ui.postMessage({
        type: 'set-line-height-result',
        updatedCount: updatedCount,
        errors: errors,
        error: errors.length > 0 ? `${errors.length}件のスタイルで更新エラーが発生しました` : null
    });
}

// 文字サイズの大きいスタイルを適用
async function applyLargerStyle() {
    const selection = figma.currentPage.selection;
    
    if (selection.length === 0) {
        return;
    }
    
    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }
    
    if (textNodes.length === 0) {
        return;
    }
    
    // すべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    if (textStyles.length === 0) {
        return;
    }
    
    let appliedCount = 0;
    
    for (const textNode of textNodes) {
        try {
            // 現在適用されているスタイルIDを取得
            let currentStyleId = textNode.textStyleId;
            
            // テキストノードのフォントファミリーを取得
            let targetFontFamily = null;
            let currentFontSize = null;
            
            // 野良テキスト（スタイル未適用）の場合
            if (!currentStyleId || currentStyleId === '' || currentStyleId === figma.mixed) {
                // テキストノードのプロパティを取得
                const textProps = getTextProperties(textNode);
                if (!textProps) {
                    continue;
                }
                targetFontFamily = textProps.fontFamily;
                currentFontSize = textProps.fontSize;
            } else {
                // 現在のスタイルを取得
                const currentStyle = textStyles.find(s => s.id === currentStyleId);
                if (!currentStyle) {
                    continue;
                }
                targetFontFamily = currentStyle.fontName.family;
                currentFontSize = currentStyle.fontSize;
            }
            
            // 同じフォントファミリーのスタイルのみをフィルタリング
            const sameFamilyStyles = textStyles.filter(style => style.fontName.family === targetFontFamily);
            
            if (sameFamilyStyles.length === 0) {
                continue;
            }
            
            // 同じフォントファミリーのスタイルをフォントサイズでソート
            const sortedStyles = sameFamilyStyles.slice().sort((a, b) => a.fontSize - b.fontSize);
            
            // 野良テキストの場合
            if (!currentStyleId || currentStyleId === '' || currentStyleId === figma.mixed) {
                // フォントサイズが最も近いスタイルを見つける
                let closestStyle = null;
                let minSizeDiff = Infinity;
                
                for (const style of sortedStyles) {
                    const sizeDiff = Math.abs(style.fontSize - currentFontSize);
                    if (sizeDiff < minSizeDiff) {
                        minSizeDiff = sizeDiff;
                        closestStyle = style;
                    }
                }
                
                if (closestStyle) {
                    await textNode.setTextStyleIdAsync(closestStyle.id);
                    appliedCount++;
                }
                continue;
            }
            
            // 次に大きいスタイルを見つける
            let nextLargerStyle = null;
            for (const style of sortedStyles) {
                if (style.fontSize > currentFontSize) {
                    nextLargerStyle = style;
                    break;
                }
            }
            
            // 次に大きいスタイルがない場合（最大サイズ）、最小サイズのスタイルを適用（ループ）
            if (!nextLargerStyle) {
                nextLargerStyle = sortedStyles[0];
            }
            
            // スタイルを適用
            await textNode.setTextStyleIdAsync(nextLargerStyle.id);
            appliedCount++;
        } catch (error) {
            console.error('Error applying larger style:', error);
        }
    }
    
    // UIが開いている場合のみスキャンを再実行して結果を更新
    if (appliedCount > 0 && isUIVisible) {
        setTimeout(() => {
            scanStyles();
        }, 500);
    }
}

// 文字サイズの小さいスタイルを適用
async function applySmallerStyle() {
    const selection = figma.currentPage.selection;
    
    if (selection.length === 0) {
        return;
    }
    
    // 対象となるテキストノードを収集
    const textNodes = [];
    for (const node of selection) {
        collectTextNodes(node, textNodes);
    }
    
    if (textNodes.length === 0) {
        return;
    }
    
    // すべてのテキストスタイルを取得
    const textStyles = await figma.getLocalTextStylesAsync();
    
    if (textStyles.length === 0) {
        return;
    }
    
    let appliedCount = 0;
    
    for (const textNode of textNodes) {
        try {
            // 現在適用されているスタイルIDを取得
            let currentStyleId = textNode.textStyleId;
            
            // テキストノードのフォントファミリーを取得
            let targetFontFamily = null;
            let currentFontSize = null;
            
            // 野良テキスト（スタイル未適用）の場合
            if (!currentStyleId || currentStyleId === '' || currentStyleId === figma.mixed) {
                // テキストノードのプロパティを取得
                const textProps = getTextProperties(textNode);
                if (!textProps) {
                    continue;
                }
                targetFontFamily = textProps.fontFamily;
                currentFontSize = textProps.fontSize;
            } else {
                // 現在のスタイルを取得
                const currentStyle = textStyles.find(s => s.id === currentStyleId);
                if (!currentStyle) {
                    continue;
                }
                targetFontFamily = currentStyle.fontName.family;
                currentFontSize = currentStyle.fontSize;
            }
            
            // 同じフォントファミリーのスタイルのみをフィルタリング
            const sameFamilyStyles = textStyles.filter(style => style.fontName.family === targetFontFamily);
            
            if (sameFamilyStyles.length === 0) {
                continue;
            }
            
            // 同じフォントファミリーのスタイルをフォントサイズでソート（降順）
            const sortedStyles = sameFamilyStyles.slice().sort((a, b) => b.fontSize - a.fontSize);
            
            // 野良テキストの場合
            if (!currentStyleId || currentStyleId === '' || currentStyleId === figma.mixed) {
                // フォントサイズが最も近いスタイルを見つける
                let closestStyle = null;
                let minSizeDiff = Infinity;
                
                for (const style of sortedStyles) {
                    const sizeDiff = Math.abs(style.fontSize - currentFontSize);
                    if (sizeDiff < minSizeDiff) {
                        minSizeDiff = sizeDiff;
                        closestStyle = style;
                    }
                }
                
                if (closestStyle) {
                    await textNode.setTextStyleIdAsync(closestStyle.id);
                    appliedCount++;
                }
                continue;
            }
            
            // 次に小さいスタイルを見つける（降順でソートされているので、最初に見つかった小さいものが次に小さい）
            let nextSmallerStyle = null;
            for (const style of sortedStyles) {
                if (style.fontSize < currentFontSize) {
                    nextSmallerStyle = style;
                    break;
                }
            }
            
            // 次に小さいスタイルがない場合（最小サイズ）、最大サイズのスタイルを適用（ループ）
            if (!nextSmallerStyle) {
                nextSmallerStyle = sortedStyles[0]; // 降順ソートなので最初が最大
            }
            
            // スタイルを適用
            await textNode.setTextStyleIdAsync(nextSmallerStyle.id);
            appliedCount++;
        } catch (error) {
            console.error('Error applying smaller style:', error);
        }
    }
    
    // UIが開いている場合のみスキャンを再実行して結果を更新
    if (appliedCount > 0 && isUIVisible) {
        setTimeout(() => {
            scanStyles();
        }, 500);
    }
}

// スタイル一覧を取得
async function getStylesList() {
    const textStyles = await figma.getLocalTextStylesAsync();
    
    const styles = textStyles.map(style => {
        // 行間の表示用文字列を生成
        let lineHeightStr = 'Auto';
        if (style.lineHeight && style.lineHeight.unit !== 'AUTO') {
            if (style.lineHeight.unit === 'PIXELS') {
                lineHeightStr = `${Math.round(style.lineHeight.value * 100) / 100}px`;
            } else if (style.lineHeight.unit === 'PERCENT') {
                lineHeightStr = `${Math.round(style.lineHeight.value)}%`;
            }
        }
        
        // レタースペーシングの表示用文字列を生成
        let letterSpacingStr = '0%';
        if (style.letterSpacing) {
            if (style.letterSpacing.unit === 'PIXELS') {
                letterSpacingStr = `${Math.round(style.letterSpacing.value * 100) / 100}px`;
            } else if (style.letterSpacing.unit === 'PERCENT') {
                letterSpacingStr = `${Math.round(style.letterSpacing.value * 10) / 10}%`;
            }
        }
        
        return {
            id: style.id,
            name: style.name,
            lineHeight: lineHeightStr,
            letterSpacing: letterSpacingStr,
            fontFamily: style.fontName.family,
            fontWeight: style.fontName.style
        };
    });
    
    figma.ui.postMessage({
        type: 'styles-list-result',
        styles: styles
    });
}

// スタイルを一括編集
async function bulkEditStyles(property, value, styleIds) {
    const textStyles = await figma.getLocalTextStylesAsync();
    
    let updatedCount = 0;
    let skippedCount = 0;
    const errors = [];
    
    for (const styleId of styleIds) {
        const style = textStyles.find(s => s.id === styleId);
        if (!style) {
            skippedCount++;
            continue;
        }
        
        try {
            // フォントを読み込む（必要な場合）
            await figma.loadFontAsync({
                family: style.fontName.family,
                style: style.fontName.style
            });
            
            if (property === 'fontSize') {
                // フォントサイズを変更
                style.fontSize = value;
                updatedCount++;
            } else if (property === 'lineHeight') {
                // 行間を変更（パーセント）
                style.lineHeight = {
                    unit: 'PERCENT',
                    value: value
                };
                updatedCount++;
            } else if (property === 'letterSpacing') {
                // レタースペースを変更（パーセント）
                style.letterSpacing = {
                    unit: 'PERCENT',
                    value: value
                };
                updatedCount++;
            } else if (property === 'weight') {
                // ウェイトを変更
                const newFontName = {
                    family: style.fontName.family,
                    style: value
                };
                
                // 新しいフォントを読み込めるかチェック
                try {
                    await figma.loadFontAsync(newFontName);
                    style.fontName = newFontName;
                    updatedCount++;
                } catch (fontError) {
                    // フォントが存在しない場合はスキップ
                    console.error(`Font not found: ${newFontName.family} ${newFontName.style}`);
                    skippedCount++;
                    errors.push({
                        name: style.name,
                        error: `${value}ウェイトが存在しません`
                    });
                }
            }
        } catch (error) {
            console.error(`Error updating style ${style.name}:`, error);
            skippedCount++;
            errors.push({
                name: style.name,
                error: error.message || 'Unknown error'
            });
        }
    }
    
    // 結果を送信
    figma.ui.postMessage({
        type: 'bulk-edit-result',
        updatedCount: updatedCount,
        skippedCount: skippedCount,
        errors: errors,
        error: errors.length > 0 ? `${errors.length}件でエラー: ${errors.map(e => e.name + ' - ' + e.error).join(', ')}` : null
    });
}

// 指定されたスタイル（複数可）が適用されているテキストノードを選択
async function selectNodesWithStyles(styleIds) {
    // ページ全体のテキストノードを収集
    const allTextNodes = [];
    collectTextNodes(figma.currentPage, allTextNodes);
    
    // スタイルIDのセットを作成（高速な検索のため）
    const styleIdSet = new Set(styleIds);
    
    // 指定されたスタイルが適用されているノードを収集
    const nodesToSelect = [];
    
    for (const textNode of allTextNodes) {
        const textStyleId = textNode.textStyleId;
        
        // 完全一致
        if (styleIdSet.has(textStyleId)) {
            nodesToSelect.push(textNode);
            continue;
        }
        
        // mixedの場合は各範囲をチェック
        if (textStyleId === figma.mixed) {
            try {
                const length = textNode.characters.length;
                let found = false;
                for (let i = 0; i < length && !found; i++) {
                    const rangeStyleId = textNode.getRangeTextStyleId(i, i + 1);
                    if (styleIdSet.has(rangeStyleId)) {
                        nodesToSelect.push(textNode);
                        found = true;
                    }
                }
            } catch (error) {
                console.error('Error checking mixed styles:', error);
            }
        }
    }
    
    // ノードを選択
    if (nodesToSelect.length > 0) {
        figma.currentPage.selection = nodesToSelect;
    }
}

