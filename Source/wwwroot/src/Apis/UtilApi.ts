export async function getGraphTokenUsingSsoToken(ssoToken: string): Promise<string> {
    return fetch(`/api/graphtoken`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + ssoToken,
        }),
    }).then((res) => {
        return res.text();
    });
}