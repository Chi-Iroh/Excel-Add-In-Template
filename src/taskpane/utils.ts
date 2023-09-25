/**
 *  Displays error on the task pane.
 * @param error the caught error
 */
export function displayErrorOnTaskpane(error : Error) : void {
    document.getElementById("error")!.innerText = `${error.stack}\n${error.name} : ${error.message}`;
}

/**
 *  Erases error on the task pane.
 */
export function eraseErrorOnTaskpane() : void {
    document.getElementById("error")!.innerText = "";
}