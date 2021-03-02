import { OpenXmlPackage } from "./open-xml-package";
import { Relationship } from "./relationship";

export class Part {
    rels: Relationship[];

    constructor(public path: string) {
    }

    load(pkg: OpenXmlPackage): Promise<any> {
        return pkg.loadRelationships(this.path).then(rels => {
            this.rels = rels;
        })
    }
}