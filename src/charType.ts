export const ALPHABET_UPPERCASE : string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
export const ALPHABET_LOWERCASE : string = ALPHABET_UPPERCASE.toLowerCase();

/**
 * Converts an ASCII code to a char (string of length 1)
 * @param code the ASCII code
 * @returns the char corresponding to the code
 */
export function fromAscii(code : number) : string {
    return String.fromCharCode(code);
}

/**
 * Converts a char to its ASCII number
 * @param char the char (if char.length > 1, char[0] is used)
 * @returns the ASCII number corresponding to the char
 * @throws Error if char is empty
 */
export function toAscii(char : string) : number {
    if (char.length == 0) {
        throw new Error("char must be a non-empty string !");
    }
    return char.charCodeAt(0);
}

/**
 * Adds a char and an ASCII offset
 * @param char a char (if char.length > 1, char[0] is used)
 * @param offset an offset (either positive or negative)
 * @returns the char
 * @throws Error if char is empty
 */
export function addAscii(char : string, offset : number) : string {
    if (char.length == 0) {
        throw new Error("char must be a non-empty string !");
    }
    return fromAscii(toAscii(char) + offset);
}

/**
 * Checks whether a char is a letter (A-Z or a-z) or not
 * @param char a char (if char.length > 1, char[0] is used)
 * @returns true if it is, false otherwise
 * @throws Error if char is empty
 */
export function isLetter(char : string) : boolean {
    if (char.length == 0) {
        throw new Error("char must be a non-empty string !");
    }
    return /^[A-Za-z]$/i.test(char[0]);
}

/**
 * Checks whether a char is a letter (0-9) or not
 * @param char a char (if char.length > 1, char[0] is used)
 * @returns true if it is, false otherwise
 * @throws Error if char is empty
 */
export function isDigit(char : string) : boolean {
    if (char.length == 0) {
        throw new Error("char must be a non-empty string !");
    }
    return /^[0-9]$/i.test(char[0]);
}