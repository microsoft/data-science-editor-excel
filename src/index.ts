import { saveSetting, loadSetting, SettingsKey } from "./xl/storage";
import { blocks, category, transforms, setCurrentWorkspace } from "./blocks";
import { getAllTables } from "./xl/tables";

interface DataScienceEditorPostPayload {
    // TODO: replace these types with the actual types
    type?: "dsl";
    action?: string;
    dslid?: string;
    options?: { table: [string, string][] };
    editor?: string;
    xml?: string;
    json?: object;
}

function post(payload: DataScienceEditorPostPayload) {
    const dseditor = document.getElementById("dseditor") as HTMLIFrameElement;
    if (!dseditor) {
        console.debug(`dseditor undefined`);
        return;
    }
    console.debug(`data blocks send`, payload);
    dseditor.contentWindow.postMessage(payload, "*");
}

async function postTables(currentDslId: string) {
    const table: [string, string][] = [];
    await Excel.run(async (context) => {
        const tables = await getAllTables(context);
        for (const t of tables) {
            if (t.name.charAt(0) !== "_") {
                table.push([t.name, t.name]);
            }
        }
    });

    post({
        type: "dsl",
        dslid: currentDslId,
        action: "options",
        options: {
            table,
        },
    });
}

async function handleBlocks(data) {
    console.debug(`hostdsl: sending blocks`);
    post({ ...data, blocks, category });
}

async function handleTransform(data) {
    const { blockId, workspace, dataset, ...rest } = data;
    let result: object;
    const block = workspace.blocks.find(({ id }) => id === blockId);
    if (!block) {
        console.error(`block ${blockId} not found in workspace`);
        result = { warning: "block lost" };
    } else {
        const transform = transforms[block.type];
        result = await transform(block, dataset);
    }
    post({ ...rest, ...(result || {}) });
}

(() => {
    // editor identifier sent by the embedded block editor
    let currentDslId;
    let loaded = false;
    let pendingLoad: { editor: string; xml: string; json: object };

    const tryLoading = async () => {
        if (!pendingLoad || !currentDslId) return;

        const { editor, xml, json } = pendingLoad;
        console.debug(`settings.sending`, { editor, xml, json });
        pendingLoad = undefined;
        await postTables(currentDslId);
        post({
            type: "dsl",
            action: "load",
            editor,
            xml,
            json,
        });
    };

    const handleMessage = (
        msg: MessageEvent<{
            // TODO: replace these types with the actual types
            type: string;
            dslid: string;
            action: string;
            workspace: string;
            editor: string;
            xml: string;
            json: string;
        }>
    ) => {
        const { data } = msg;
        if (data.type !== "dsl") {
            return;
        }
        const { dslid, action } = data;
        console.debug(action, data);
        switch (action) {
            case "mount": {
                currentDslId = dslid;
                console.debug(`dslid: ${dslid}`);
                tryLoading();
                break;
            }
            case "unmount": {
                currentDslId = undefined;
                break;
            }
            case "blocks": {
                handleBlocks(data);
                break;
            }
            case "transform": {
                handleTransform(data);
                break;
            }
            case "workspace": {
                const { workspace } = data;
                setCurrentWorkspace(workspace);
                break;
            }
            case "save": {
                // don't save until we've reloaded our content from excel
                if (!loaded) {
                    console.debug(`save.ignore: not loaded yet`);
                    break;
                }

                const { editor, xml, json } = data;
                const file = {
                    editor,
                    xml,
                    json,
                };
                saveSetting(SettingsKey.EditorSaveData, JSON.stringify(file));
                break;
            }
            case "change": {
                handleTransform(data);
                break;
            }
        }
    };

    Office.onReady(() => {
        loadSetting(SettingsKey.EditorSaveData).then((setting) => {
            loaded = true;
            if (!setting) {
                console.debug(`settings.none`);
                return;
            }

            const parsed = JSON.parse(setting);
            pendingLoad = parsed;
            console.debug(`settings.found`, { toLoad: pendingLoad, setting });
            tryLoading();
        });

        window.addEventListener("message", handleMessage, false);

        Excel.run(async (context) => {
            console.debug(`dsl: initializing`);
            context.workbook.tables.onChanged.add(onTableChanged);
            await context.sync();
            console.debug(`dsl: initialized`);
        });
    });

    async function onTableChanged(eventArgs: Excel.TableChangedEventArgs) {
        await postTables(currentDslId);
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const table = sheet.tables.getItem(eventArgs.tableId);
            sheet.load("id");
            table.load("name");
            await context.sync();

            // Only track changes to tables on the active sheet
            if (
                currentDslId &&
                eventArgs.worksheetId === sheet.id &&
                table.name.charAt(0) !== "_"
            ) {
                post({
                    type: "dsl",
                    dslid: currentDslId,
                    action: "change",
                });
            }
        });
    }
})();
