export function deobfuscate(data: number[], guidKey: string): Promise<string> {
    const hexStrings = guidKey.replace(/{|}|-/g, "").replace(/(..)/g, "$1 ").trim().split(" ")
    const hexNumbers = hexStrings.map((hexString) => parseInt(hexString, 16))
    hexNumbers.reverse()

    var array = new Uint8Array(data);

    for (let i = 0; i < 32; i++) {
        array[i] = array[i] ^ hexNumbers[i % hexNumbers.length]
    }

    return new Promise(resolve => {
        var reader = new FileReader();
        reader.onload = (event) => resolve(event.target.result as string);
        reader.readAsDataURL(new Blob([array]));
    });
}