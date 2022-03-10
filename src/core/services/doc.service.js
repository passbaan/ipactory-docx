/* eslint-disable no-case-declarations */
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  hexColorValue,
  TableCell,
  Table,
  TableRow,
  NumberFormat,
  MathRun,
  Math,
  Header,
  WidthType,
  MathSuperScript,
  MathRadical,
  MathFraction,
  MathSubScript,
  PageNumber,
  VerticalAlign,
  LineNumberRestartFormat,
  BorderStyle,
  ImageRun,
  PageBreak,
  HeadingLevel,
} from "docx";
import fig1 from "/public_assets/images/fig1.jpg";
import fig2 from "/public_assets/images/fig2.jpg";
import fig3 from "/public_assets/images/fig3.jpg";
import fig4 from "/public_assets/images/fig4.jpg";
const images = [fig1, fig2, fig3, fig4];
// console.log("file: do363c.service.js | line 29 | images", images);
import * as _ from "lodash";
import { saveAs } from "file-saver";
const calcTWIP = (val) => {
  if (val.includes("cm")) {
    return parseInt(val.replace("cm")) * 566.9291338583;
  } else if (val.includes("pt")) {
    return parseInt(val.replace("pt")) * 20;
  }
  return 0;
};

const STYLES = [
  {
    id: "alignment",
    key: "text-align",
    isStatic: true,
    values: {
      center: AlignmentType.CENTER,
      left: AlignmentType.LEFT,
      right: AlignmentType.RIGHT,
      justify: AlignmentType.JUSTIFIED,
    },
  },
  {
    id: "color",
    key: "color",
    isStatic: false,
    value: (v) => hexColorValue(v),
  },
  {
    id: "font",
    key: "font-family",
    isStatic: false,
    value: (v) => v,
  },
  {
    id: "size",
    key: "font-size",
    isStatic: false,
    value: (v) => parseFloat(v.replace("pt", "")),
  },
  {
    id: "indent",
    key: "padding",
    isStatic: false,
    value: (val) => {
      const vals = val.split(" ");

      return {
        top: calcTWIP(vals[0]),
        right: calcTWIP(vals[1]),
        bottom: calcTWIP(vals[2]),
        left: calcTWIP(vals[3]),
      };
    },
  },
  {
    id: "highlight",
    key: "background-color",
    isStatic: false,
    value: (v) => hexColorValue(v),
  },
  {
    id: "width",
    key: "width",
    isStatic: false,
    value: (v) => ({
      size: calcTWIP(v),
      type: WidthType.DXA,
    }),
  },
  {
    id: "spacing",
    key: "line-height",
    isStatic: false,
    value: (v) => ({
      line:
        typeof v === "string" ? 240 : parseFloat(((v * 240) / 100).toFixed(2)),
    }),
  },
];
const getStyles = (key, value) => {
  const style = STYLES.find((style) => style.key === key);
  if (style) {
    return [
      style.id,
      style.isStatic && style.values[value]
        ? style.values[value]
        : style.value(value),
    ];
  }
  return null;
};
const attrStringToJson = (string) => {
  let xx = string.split(";"),
    i = xx.length,
    json = { style: {} },
    style,
    k,
    v;

  while (i--) {
    style = xx[i].split(":");
    k = _.trim(style[0]);
    v = _.trim(style[1]);
    if (k.length > 0 && v.length > 0) {
      const styleItem = getStyles(k, v);
      if (styleItem !== null) {
        json.style[styleItem[0]] = styleItem[1];
      } else {
        json.style[k] = v;
      }
    }
  }
  return json.style;
};

const preserverParentStyle = (item) => {
  let { state } = item;

  (function findMoreSpans(stateSpan) {
    let firstSpanIdx = stateSpan.findIndex((x) => x.key === "span");
    if (firstSpanIdx === -1) {
      return;
    }
    let ssa = stateSpan.slice(firstSpanIdx);
    if (!item.styles) {
      item.styles = { current: {} };
    }

    Object.assign(item.styles.current, ssa[0].data);
    // console.log("file: doc.service.js | line 189 | findMoreSpans | ss", ss);
    findMoreSpans(ssa.slice(1));
  })(state);
};
const traverse = (input, styles = {}, state = [], level = 0) => {
  let text = {
    type: input.type === "node" ? input.tagName : input.type,
    styles: { current: {} },
    level,
  };

  let passStyles = {};
  if (input.attrs) {
    text.attributes = input.attrs;
    if (text.attributes.style) {
      text.styles["current"] = attrStringToJson(text.attributes.style);
      passStyles = text.styles.current;
    }
  }
  let x = {};
  x["key"] = text.type;
  x["data"] = passStyles;
  const currentState = [...state, x];

  if (input.children) {
    text.children = input.children.map((item) => {
      return traverse(item, passStyles, currentState, level + 1);
    });
  }

  switch (text.type) {
    case "text":
      text.state = currentState;

      preserverParentStyle(text);
      text.value = input.text;
      // text.styles.current = styles;
      break;
    case "strong":
      text.children[0]["isStrong"] = true;
      text = text.children[0];
      Object.assign(text.styles.current, styles);

      break;
    case "span":
      let temp = [];
      let indexFactor = 0;
      text.children = text.children.filter((ch, i) => {
        if (ch.children) {
          ch.children.forEach((c) => temp.push({ i, c }));
          return false;
        }
        return true;
      });
      if (temp.length > 0) {
        temp.forEach((x) => {
          text.children.splice(x.i + indexFactor, 0, x.c);
          indexFactor += 1;
        });
      }
      break;
    case "sub":
      // preserverParentStyle(text.children[0]);
      text.children[0]["subScript"] = true;
      break;
    case "u":
      // preserverParentStyle(text.children[0]);
      text.children[0]["underline"] = true;
      break;
    case "s":
      // preserverParentStyle(text.children[0]);
      text.children[0]["strike"] = true;
      break;

    default:
      text.__v = 1;
  }
  text.state = currentState;
  return text;
};

const generate = (x, count = { p: 0 }) => {
  let children = null;
  if (x.children) {
    children = x.children
      .map((item) => {
        return generate(item, count);
      })
      .filter((i) => i !== null);
  }
  if (x.type === "root") {
    return {
      section: children,
    };
  } else if (x.type === "p") {
    const test = Object.entries(x.styles)[0];
    let style = {};
    if (test !== null) {
      style = test[1];
    }
    let newChildren = [];
    children.forEach((c) => {
      if (Array.isArray(c)) {
        newChildren = [...newChildren, ...c];
      } else {
        newChildren.push(c);
      }
    });
    if (x.level === 1) {
      count.p += 1;
      if (count.p !== 1) {
        newChildren.unshift(
          new TextRun({ text: `[${pad(count.p - 1, 4)}]  ` })
        );
      }
      if (count.p === 1) {
        style.heading = HeadingLevel.HEADING_1;
      }
    }
    if ("size" in style) {
      style.size *= 2;
    }

    return new Paragraph({
      children: newChildren,
      ...style,
    });
  } else if (x.type === "text") {
    const text = x;
    if (text.value === " ") {
      return null;
    }
    const test = Object.entries(x.styles)[0];
    let current = {};
    if (test !== null) {
      current = test[1];
    }
    if (!current) {
      current = {};
    }
    current.bold = !!text.isStrong;
    current.subScript = !!text.subScript;
    if (text.underline) {
      current.underline = {};
    }
    if (text.strike) {
      current.strike = {};
    }
    if ("size" in current) {
      current.size *= 2;
    }

    return new TextRun({
      text: text.value,
      ...current,
    });
  } else if (x.type === "span") {
    return children;
    // console.log("file: doc.service.js | line 159 | generate | x", x);
  } else if (x.type === "strong" || x.type === "mfenced") {
    return children[0];
  } else if (x.type === "sub") {
    return children[0];
  } else if (x.type === "u") {
    return children[0];
  } else if (x.type === "s") {
    return children[0];
  } else if (x.type === "img") {
    return new TextRun({
      text: "(---Image here---)",
    });
  } else if (x.type === "comment") {
    return new PageBreak();
  } else if (x.type === "table") {
    return children[0];
  } else if (x.type === "div") {
    return children;
  } else if (x.type === "br") {
    return new TextRun({
      text: "break",
      break: 1,
    });
  } else if (x.type === "tbody") {
    return new Table({
      rows: children,
    });
  } else if (x.type === "tr") {
    return new TableRow({
      children,
    });
  } else if (x.type === "td") {
    const localSettings = {};
    Object.entries(x.attributes).forEach(([k, v]) => {
      if (k === "colspan") {
        localSettings["columnSpan"] = v;
      }
    });
    return new TableCell({
      children,
      ...localSettings,
      ...x.styles.current,
    });
  } else if (x.type === "math") {
    return new Math({
      children,
    });
  } else if (x.type === "mi" || x.type === "mo" || x.type === "mn") {
    return new MathRun(x.children[0].value);
  } else if (x.type === "msqrt") {
    return new MathRadical({
      children,
    });
  } else if (x.type === "msup") {
    return new MathSuperScript({
      children: [children[0]],
      superScript: [children[1]],
    });
  } else if (x.type === "msub") {
    return new MathSubScript({
      children: [children[0]],
      subScript: [children[1]],
    });
  } else if (x.type === "mrow") {
    return children;
  } else if (x.type === "mfrac") {
    return new MathFraction({
      numerator: children[0],
      denominator: children[1],
    });
  }

  return null;
};
const fileToDataUri = (file) =>
  new Promise((resolve) => {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", file, true);
    xhr.responseType = "blob";
    xhr.onload = function () {
      console.log(this.response);
      var reader = new FileReader();
      reader.onload = function (event) {
        var res = event.target.result;
        resolve(res);
      };
      var file = this.response;
      reader.readAsDataURL(file);
    };
    xhr.send();
  });

function pad(num, size) {
  num = num.toString();
  while (num.length < size) num = "0" + num;
  return num;
}
const createNew = async (json) => {
  // const { children } = json;
  const x = traverse(json);
  let generated = generate(x).section.filter((i) => i);
  if (images) {
    const resolved = await Promise.all(
      Object.entries(images).map(async (image) => {
        const response = {};
        response[image[0]] = await fileToDataUri(image[1]);
        return response;
      })
    );
    resolved.forEach((item) => {
      const data = Object.entries(item)[0];
      const newLabel = new Paragraph({
        children: [new TextRun({ text: data[0] })],
        alignment: AlignmentType.CENTER,
      });
      generated.push(newLabel);
      const newImage = new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: data[1],
            transformation: {
              width: 500,
              height: 500,
            },
          }),
          new PageBreak(),
        ],
        spacing: {
          before: 200,
        },
      });
      generated.push(newImage);
    });
  }
  const doc = new Document({
    sections: [
      {
        properties: {
          lineNumbers: {
            countBy: 5,
            restart: LineNumberRestartFormat.NEW_PAGE,
            size: 10,
          },
          page: {
            margin: {
              top: 1500,
              right: 1500,
              bottom: 1500,
              left: 1500,
            },
            pageNumbers: {
              start: 1,
              formatType: NumberFormat.DECIMAL,
            },
            size: {
              height: 16839,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Table({
                alignment: AlignmentType.CENTER,
                width: {
                  size: 100,
                  type: WidthType.PERCENTAGE,
                },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        columnSpan: 1,
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: "IPACTORY", size: 24 }),
                            ],
                            alignment: AlignmentType.LEFT,
                          }),
                        ],
                        width: {
                          size: 33,
                          type: WidthType.PERCENTAGE,
                        },
                        borders: {
                          top: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          bottom: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          left: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          right: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                        },
                      }),
                      new TableCell({
                        columnSpan: 1,
                        children: [
                          new Paragraph({
                            size: 12,
                            font: "Times New Roman",
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                children: [
                                  new TextRun({
                                    text: "- ",
                                    size: 24,
                                    font: "Times New Roman",
                                  }),
                                  new TextRun({
                                    children: [PageNumber.CURRENT],
                                    font: "Times New Roman",
                                    size: 24,
                                  }),
                                  new TextRun({
                                    text: " -",
                                    size: 24,
                                    font: "Times New Roman",
                                  }),
                                ],
                              }),
                            ],
                          }),
                        ],

                        verticalAlign: VerticalAlign.CENTER,
                        borders: {
                          top: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          bottom: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          left: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          right: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                        },
                        width: {
                          size: 33,
                          type: WidthType.PERCENTAGE,
                        },
                      }),
                      new TableCell({
                        columnSpan: 1,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [
                              new TextRun({ text: "259889523", size: 24 }),
                            ],
                          }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                        borders: {
                          top: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          bottom: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          left: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                          right: {
                            style: BorderStyle.NONE,
                            size: 0,
                            color: "FFFFFF",
                          },
                        },
                        width: {
                          size: 33,
                          type: WidthType.PERCENTAGE,
                        },
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        },
        children: [...generated],
      },
    ],
  });
  // Used to export the file into a .docx file
  const mimeType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  return new Promise((resolve, reject) => {
    return Packer.toBlob(doc)
      .then((blob) => {
        // console.log("file: doc.service.js | line 36 | .then | blob", blob);
        const docblob = blob.slice(0, blob.size, mimeType);
        saveAs(docblob, "output.docx");
        resolve(docblob);
      })
      .catch((error) => {
        console.error(error);
        reject(error);
      });
  });

  // Done! A file called 'My Document.docx' will be in your file system.
};
export default createNew;
