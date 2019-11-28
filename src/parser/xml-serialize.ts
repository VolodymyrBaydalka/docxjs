const schemaSymbol = Symbol("open-xml-schema");

export type Converter = (val: string) => any;

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

export function fromText(convert: Converter = null) {
    return function (target: any, prop: string) {
        var schema = getPrototypeXmlSchema(target);
        schema.text = { prop, convert };
    }
}

export function fromAttribute(attrName: string, convert: Converter = null) {
    return function (target: any, prop: string) {
        var schema = getPrototypeXmlSchema(target);
        schema.attrs[attrName] = { prop, convert };
    }
}

export function deserialize(n: Element, output: any) {
    var proto = Object.getPrototypeOf(output);
    var schema = proto[schemaSymbol];

    if (schema == null)
        return output;

    if (schema.text) {
        let prop = schema.text;
        output[prop.prop] = prop.convert ? prop.convert(n.textContent) : n.textContent; 
    }

    for (let i = 0, l = n.attributes.length; i < l; i++) {
        let attr = n.attributes.item(i);
        let prop = schema.attrs[attr.localName];

        if(prop == null)
            continue;

        output[prop.prop] = prop.convert ? prop.convert(attr.value) : attr.value; 
    }

    for (let i = 0, l = n.children.length; i < l; i ++) {
        let elem = n.children.item(i);
        let child = schema.children[elem.localName];

        if (child) {
            let obj = Object.create(child.proto);
            deserialize(elem, obj);
            output.children.push(obj);
        }
    }

    return output;
}

interface OpenXmlSchema {
    elemName: string;
    text: OpenXmlSchemaProperty;
    attrs: Record<string, OpenXmlSchemaProperty>;
    children: Record<string, any>;
}

interface OpenXmlSchemaProperty {
    prop: string;
    convert: Converter;
}

function getPrototypeXmlSchema(proto: any): OpenXmlSchema {
    return proto[schemaSymbol] || (proto[schemaSymbol] = {
        text: null,
        attrs: {},
        children: {}
    });
}