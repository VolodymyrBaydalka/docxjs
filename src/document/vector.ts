import { OpenXmlElement } from "./dom";

export interface VmlShape extends OpenXmlElement {
	cssStyleText: string;
	imagedata: {
		id: string,
		title: string
	}
}
