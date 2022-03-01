/* global Office */
export const logObject = (obj: Record<string, unknown>): void => {
  const props = Object.getOwnPropertyNames(obj);
  const objStringified = props.reduce((acc, prop) => {
    acc += `${prop}: ${String(obj[prop])}\n`;
    return acc;
  }, "");

  log(objStringified);
};

export function log(text) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.prependAsync(
      `${text}\n`,
      {
        coercionType: Office.CoercionType.Text,
      },
      (asyncResult) => {
        resolve(asyncResult.value);
      }
    );
  });
}
