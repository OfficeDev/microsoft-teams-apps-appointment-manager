export async function getSettings(): Promise<Settings> {
    return fetch("/api/settings", {
        method: "GET",
    }).then(async (response) => {
        if (!response.ok) {
            throw new Error(`Unable to get the app's settings`);
        }

        return response.json() as Promise<Settings>;
    });
}

interface Settings {
    defaultLocale: string;
}