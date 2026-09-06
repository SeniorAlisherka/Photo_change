declare module "libheif-js/libheif-wasm/libheif-bundle.mjs" {
  type HeifImage = {
    get_width(): number;
    get_height(): number;
    display(image: ImageData, callback: (data: ImageData | null) => void): void;
    free(): void;
  };

  type HeifModule = {
    HeifDecoder: new () => { decode(data: Uint8Array): HeifImage[] };
  };

  export default function createDecoder(): Promise<HeifModule>;
}
