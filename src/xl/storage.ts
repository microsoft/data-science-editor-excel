export enum SettingsKey {
    EditorSaveData = "EditorSaveData",
}

export async function saveSetting(key: SettingsKey, value: string): Promise<void> {
    await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        settings.add(SettingsKey[key], value);

        console.debug(`settings.save`, { key, value });
        await context.sync();
    });
}

export async function loadSetting(key: SettingsKey): Promise<string> {
    return await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const setting = settings.getItemOrNullObject(SettingsKey[key]);

        await context.sync();

        console.debug("settings.object", { settings });
        if (setting.isNullObject) {
            console.debug(`settings.notfound`);
            return "";
        }

        setting.load("value");
        await context.sync();

        const v = typeof setting.value === "string" ? setting.value : "";
        console.debug("settings.value", { settings, v });
        return v;
    });
}
