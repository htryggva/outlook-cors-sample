export function callApi(url: string, requestInit?: RequestInit): Promise<string> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();

    xhr.onload = function () {
      resolve(xhr.responseText);
    };

    xhr.onerror = function () {
      const error = {
        status: xhr.status,
        statusText: xhr.statusText,
        responseText: xhr.responseText,
      };
      reject(error);
    };

    const method = requestInit?.method ?? "GET";

    const headers = requestInit?.headers as Record<string, string> | undefined;

    const body = requestInit?.body as string | undefined;

    xhr.open(method, url, true);

    if (headers) {
      for (const header of Object.keys(headers)) {
        const val = headers[header];
        xhr.setRequestHeader(header, val);
      }
    }

    xhr.send(body);
  });
}
