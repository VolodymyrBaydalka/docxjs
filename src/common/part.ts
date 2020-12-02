import { Package } from "./package";
import { Relationship } from "./relationship";

export class Part {
    rels: Relationship[];

    constructor(public path: string) {
    }

    load(pkg: Package): Promise<any> {
        return pkg.loadRelationships(this.path).then(rels => {
            this.rels = rels;
        })
    }
}