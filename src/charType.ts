export const ALPHABET_UPPERCASE : string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
export const ALPHABET_LOWERCASE : string = ALPHABET_UPPERCASE.toLowerCase();

export function fromAscii(code : number) : string {
    return String.fromCharCode(code);
}

export function toAscii(char : string) : number {
    if (char.length != 1) {
        throw new Error("Argument must be a single character !");
    }
    return char.charCodeAt(0);
}

export function addAscii(char : string, offset : number) : string {
    return fromAscii(toAscii(char) + offset);
}

export function isLetter(char : string) : boolean {
    if (typeof char !== "string" || char.length != 1) {
        return false;
    }
    return /^[A-Za-z]$/i.test(char);
}

export function isDigit(char : string) : boolean {
    if (typeof char !== "string" || char.length != 1) {
        return false;
    }
    return /^[0-9]$/i.test(char);
}
export function isMathOperator(char : string) : boolean {
    if (typeof char !== "string" || char.length != 1) {
        return false;
    }
    return "+-*/^%".includes(char);
}