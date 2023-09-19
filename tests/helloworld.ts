import { hello } from "../src/helloworld";

describe("Test 1", () => {
  test("Hello, World", () => {
    expect(hello()).toBe("Hello, World !");
  });
})