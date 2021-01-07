export async function getUserPhoto(token: string, userId: string): Promise<Blob> {
    return fetch(`https://graph.microsoft.com/v1.0/users/${userId}/photos/48x48/$value`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
            "Content-Type": "image/jpg",
        }),
    }).then(response => {
        if (!response.ok) {
            throw Error('failed to get photo from Graph');
        }

        return response.blob();
    });
}

export async function getBookingsBusinesses(token: string): Promise<BookingsBusiness[]> {
    return fetch("https://graph.microsoft.com/beta/bookingBusinesses", {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response =>
        response.json()
    ).then((jsonResponse: { value: BookingsBusiness[] }) => {
        return jsonResponse.value;
    });
}

export async function getBookingsServices(token: string, id: string): Promise<BookingsService[]> {
    return fetch(`https://graph.microsoft.com/beta/bookingBusinesses/${id}/services`, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + token,
        }),
    }).then(response =>
        response.json()
    ).then((jsonResponse: { value: BookingsService[] }) => {
        return jsonResponse.value;
    });
}

export interface BookingsBusiness {
    id: string;
    displayName: string;
}

export interface BookingsService {
    id: string;
    displayName: string;
}