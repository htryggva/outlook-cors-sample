/* global console */
export const logObject = (obj: Record<string, unknown>): void => {
  const props = Object.getOwnPropertyNames(obj);
  const objStringified = props.reduce((acc, prop) => {
    acc += `${prop}: ${String(obj[prop])}\n`;
    return acc;
  }, "");

  log(objStringified);
};

export function log(text: string) {
  console.log(`[outlook-cors-sample] ${text}`);
}
