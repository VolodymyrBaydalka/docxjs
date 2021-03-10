const schemaSymbol = Symbol("open-xml-schema");

export type ValueConverter = (val: string) => any;

export type ElementConverter = (val: Element) => any;

export function element(name: string) {
    return function(target: any) {
        var schema = getPrototypeXmlSchema(target.prototype);
        schema.elemName = name;
    }
}

export function children(...elements: any[]) {
    return function(target) {
        var schema = getPrototypeXmlSchema(target.prototype);
        schema.children = {};
        for(let c of elements) {
            let cs = getPrototypeXmlSchema(c.prototype);
            schema.children[cs.elemName] = { proto: c.prototype, schema: cs };
        }
    }
}

export function fromText(convert: ValueConverter = null) {
    return function (target: any, prop: string) {
        var schema = getPrototypeXmlSchema(target);
        schema.text = { prop, convert };
    }
}

export function fromAttribute(attrName: string, convert: ValueConverter = null) {
    return function (target: any, prop: string) {
        var schema = getPrototypeXmlSchema(target);
        schema.attrs[attrName] = { prop, convert };
    }
}

export function fromElement(elemName: string, convert: ElementConverter) {
    return function (target: any, prop: string) {
        var schema = getPrototypeXmlSchema(target);
        schema.elements[elemName] = { prop, convert };
    }   
}

export function buildXmlSchema(schemaObj: any): OpenXmlSchema {
    var schema: OpenXmlSchema = {
        text: null,
        attrs: {},
        elements: {},
        elemName: null,
        children: null
    };

    for(let p in schemaObj) {
        let v = schemaObj[p];

        if(p == "$elem") {
            schema.elemName = v;
        }
        else if(v.$attr) {
            schema.attrs[v.$attr] = { prop: p, convert: null };
        }
    }

    return schema;
}

export function deserializeElement<T = any>(n: Element, output: T, ops: DeserializeOptions): T {
    var proto = Object.getPrototypeOf(output);
    var schema = proto[schemaSymbol];

    if (ops?.keepOrigin) {
        (output as any).$$xmlElement = n;
    }  

    if (schema == null)
        return output;

    deserializeSchema(n, output, schema);

    for (let i = 0, l = n.children.length; i < l; i ++) {
        let elem = n.children.item(i);
        let child = schema.children[elem.localName];

        if (child) {
            let obj = Object.create(child.proto);
            deserializeElement(elem, obj, ops);
            (output as any).children.push(obj);
        }
    }

    return output;
}

export function deserializeSchema(n: Element, output: any, schema: OpenXmlSchema) {
    if (schema.text) {
        let prop = schema.text;
        output[prop.prop] = prop.convert ? prop.convert(n.textContent) : n.textContent; 
    }

    for (let i = 0, l = n.attributes.length; i < l; i++) {
        const attr = n.attributes.item(i);
        const prop = schema.attrs[attr.localName];

        if(prop == null)
            continue;

        output[prop.prop] = prop.convert ? prop.convert(attr.value) : attr.value; 
    }

    for (let i = 0, l = n.childNodes.length; i < l; i ++) {
        const elem = n.childNodes.item(i) as Element;
        const prop = elem.nodeType === Node.ELEMENT_NODE ? schema.elements[elem.localName] : null;

        if (prop == null)
            continue;

        output[prop.prop] = prop.convert(elem); 
    }

    return output;
}

export interface DeserializeOptions {
    keepOrigin: boolean
}

export interface OpenXmlSchema {
    elemName: string;
    text: OpenXmlSchemaProperty;
    attrs: Record<string, OpenXmlSchemaProperty>;
    elements: Record<string, any>;
    children: Record<string, any>;
}

export interface OpenXmlSchemaProperty {
    prop: string;
    convert: ValueConverter;
}

function getPrototypeXmlSchema(proto: any): OpenXmlSchema {
    return proto[schemaSymbol] || (proto[schemaSymbol] = {
        text: null,
        attrs: {},
        children: {},
        elements: {}
    });
}