// Promise wrappers over fflate's asynchronous zip. The async function runs the deflate
// on a Web Worker (fflate spins one up internally), so compressing a document on save no
// longer blocks the main thread. The synchronous zipSync stays for tests and the small
// fixed-size payloads (blank templates) where a worker is not worth its startup cost.
import { zip, type AsyncZippable } from "fflate";

export function zipAsync(files: AsyncZippable): Promise<Uint8Array> {
  return new Promise((resolve, reject) => zip(files, (err, data) => (err ? reject(err) : resolve(data))));
}
