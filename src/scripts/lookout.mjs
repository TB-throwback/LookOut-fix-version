import { tnef_parse } from "/scripts/tnef.mjs"

// Partially implements nsIInputStream.
// https://udn.realityripple.com/docs/Mozilla/Tech/XPCOM/Reference/Interface/nsIInputStream
class PseudoInputStream {
  constructor() {
    this.file = null;
    this.buffer = null;
    this.view = null;
    this.offset = 0;
  }

  async setFile(file) {
    this.file = file;
    this.buffer = await this.file.arrayBuffer();
    this.view = new Uint8Array(this.buffer);
    this.offset = 0;
  }

  available() {
    return this.view.length - this.offset;
  }

  test(bytes) {
    if (this.available < bytes) {
      throw new Error("Trying to read beyond the end of the arrayBuffer");
    }
  }

  readByteArray(bytes) {
    this.test(bytes);
    let byteArray = [];
    for (let i = 0; i < bytes; i++) {
      byteArray[i] = this.view[this.offset + i];
    }
    this.offset += bytes;
    return byteArray;
  }

  read8() {
    this.test(1);
    let rv = this.view[this.offset];
    this.offset += 1;
    return rv
  }

  readBytes(bytes) {
    let binaryString = "";
    let byteArray = this.readByteArray(bytes);
    for (let byte of byteArray) {
      binaryString += `${String.fromCharCode(byte)}`;
    }
    return binaryString;
  }

  close() {
    //NOOP
  }
}

export class TnefExtractor {
  constructor() {
    this.mStream = new PseudoInputStream();
    this.files = [];
    this.mPartId = 0;
  }

  async parse(file, msgHdr, prefs) {
    this.mMsgHdr = msgHdr; // TODO: Why does it need it?
    
    // The TNEF parser uses debug_level.
    prefs["debug_level"] = prefs["debug_enabled"] ? 10 : 5;

    await this.mStream.setFile(file);
    tnef_parse(
      this.mStream, 
      this.mMsgHdr, // TODO: Why does it need it?
      this,
      prefs
    );
    return this.files;
  }

  onTnefFile(data, filename, content_type, length, date) {
    // Strip away path.
    filename = filename.split('\\').pop().split('/').pop();

    if (!content_type) {
      content_type = "application/binary";
    }
  
    // The data is a binary string, but we need an Uint8Array to not trigger utf8
    // interpretation.
    let bytes = new Array(data.length);
    for (let i = 0; i < bytes.length; i++) {
      bytes[i] = data.charCodeAt(i) & 0xFF;
    }
    this.files.push(new File([new Uint8Array(bytes)], filename, {type: content_type}));
    this.mPartId++;
  }
}
