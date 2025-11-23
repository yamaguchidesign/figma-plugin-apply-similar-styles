// スタイルが当たっていないテキストを見つけて、近いスタイルを適用するプラグイン

figma.showUI(__html__, { width: 320, height: 700 });

// 選択状態の変更を監視
figma.on('selectionchange', async () => {
    await scanStyles();
});

// 初期状態をチェック
scanStyles();

figma.ui.onmessage = async (msg) => {
    if (msg.type === 'scan-styles') {
        await scanStyles();
    } else if (msg.type === 'apply-styles') {
        await applyNearestStyles();
    } else if (msg.type === 'check-unused-styles') {
        await checkUnusedStyles();
    } else if (msg.type === 'delete-unused-styles') {
        await deleteUnusedStyles();
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
    let hasNearestCount = 0; // 近いスタイルがある
    let noNearestCount = 0; // 近いスタイルがない

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
            noNearestCount++;
            continue;
        }

        // 近いスタイルを探す
        const nearestStyle = await findNearestStyle(textProps, textStyles);

        if (nearestStyle) {
            hasNearestCount++;
        } else {
            noNearestCount++;
        }
    }

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

async function applyNearestStyles() {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
        figma.ui.postMessage({
            type: 'result',
            found: 0,
            applied: 0,
            skipped: 0
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

    let foundCount = 0;
    let appliedCount = 0;
    let skippedCount = 0;
    let createdCount = 0;

    for (const textNode of textNodes) {
        // スタイルが既に当たっているかチェック
        if (hasTextStyle(textNode)) {
            continue; // スタイルが当たっている場合はスキップ
        }

        foundCount++;

        // テキストノードのプロパティを取得
        const textProps = getTextProperties(textNode);

        if (!textProps) {
            skippedCount++;
            continue;
        }

        // 近いスタイルを探す
        let nearestStyle = await findNearestStyle(textProps, textStyles);

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
            // 近いスタイルがない場合は新規作成
            try {
                const newStyle = await createTextStyleFromNode(textNode, textProps);
                if (newStyle) {
                    textStyles.push(newStyle); // 作成したスタイルをリストに追加
                    await textNode.setTextStyleIdAsync(newStyle.id);
                    createdCount++;
                    appliedCount++;
                }
            } catch (error) {
                console.error('Error creating text style:', error);
                skippedCount++;
            }
        }
    }

    // 結果を表示
    figma.ui.postMessage({
        type: 'result',
        found: foundCount,
        applied: appliedCount,
        skipped: skippedCount,
        created: createdCount
    });
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

        // フォントウェイト（style）が完全一致するかチェック
        if (styleFontName.style !== textProps.fontWeight) {
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
