import { getUserPhoto } from "../Apis/GraphApi";

export class PhotoUtil {
    // keeps a record of already loaded profile photos
    pics: Record<string, string> = {};

    // empty image to default to
    emptyPic = "/images/default_person.png";

    // gets a photo from microsoft graph for a specific user
    public getGraphPhoto = (token: string, id: string): Promise<string> => {
        // check if already loaded
        if (this.pics[id]) {
            return Promise.resolve(this.pics[id]);
        }

        return getUserPhoto(token, id).then((blob) => {
            // generate a blob url
            const url = window.URL || window.webkitURL;
            const objectURL = url.createObjectURL(blob);
            this.pics[id] = objectURL;
            return objectURL;
        }).catch((err) => {
            // revert to the empty pic
            this.pics[id] = this.emptyPic;
            return this.emptyPic;
        });
    };
}