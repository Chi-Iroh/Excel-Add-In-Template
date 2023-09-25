/**
 * Finds the first character which validates a predicate.
 * @param str the string
 * @param predicate the predicate, must take a char (string of lentgh 1) as parameter and return a boolean
 * @returns -1 if no character satisfies the condition, otherwise its index
 */
export function strFindIf(str : string, predicate: (char : string) => boolean) : number {
    for (let i = 0; i < str.length; i++) {
        if (predicate(str[i])) {
            return i;
        }
    }
    return -1;
}

/**
 * Finds the first character which doesn't validate a predicate.
 * @param str the string
 * @param predicate the predicate, must take a char (string of lentgh 1) as parameter and return a boolean
 * @returns -1 if all character satisfy the condition, otherwise its index
 */
export function strFindIfNot(str : string, predicate: (char : string) => boolean) : number {
    for (let i = 0; i < str.length; i++) {
        if (!predicate(str[i])) {
            return i;
        }
    }
    return -1;
}