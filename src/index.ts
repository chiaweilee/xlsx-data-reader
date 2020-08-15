import { read } from "./helper/xlsx";

const XLSXReader = (data: any, options?: object) => {
  const getDataType = (): {
    type?: "array" | "binary" | "base64" | "string";
  } => {
    if (typeof data === "string") {
      return { type: "array" };
    } else if (data instanceof ArrayBuffer) {
      return { type: "array" };
    }
    return {};
  };
  return read(data, { ...getDataType(), ...options });
};

export default XLSXReader;

export const xlsxHandler = callback => e => {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = (ex: any) => {
    const data = ex.target.result;
    try {
      const book = XLSXReader(data);
      if (typeof callback === "function") {
        callback(book);
      }
    } catch (e) {
      console.error(e);
    }
  };
  reader.readAsArrayBuffer(file);
};
