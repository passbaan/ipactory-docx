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
} from "docx";
import * as _ from "lodash";
// import { saveAs } from "file-saver";
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
const traverse = (input, styles = {}, state = []) => {
  let text = {
    type: input.type === "node" ? input.tagName : input.type,
    styles: { current: null },
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
  x[text.type] = passStyles;
  const currentState = [...state, x];

  if (input.children) {
    text.children = input.children.map((item) => {
      return traverse(item, passStyles, currentState);
    });
  }

  switch (text.type) {
    case "text":
      text.value = input.text;
      text.styles.current = styles;
      break;
    case "strong":
      text.children[0]["isStrong"] = true;
      text = text.children[0];
      Object.assign(text.styles.current, styles);
      break;
    case "span":
      text.styles.current = styles;
      let temp = [];
      let indexFactor = 0;
      text.children = text.children.filter((ch, i) => {
        if (ch.children) {
          console.log(
            "file: doc.service.js | line 148 | text.children=text.children.filter | ch",
            ch
          );
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
      text.children[0]["subScript"] = true;
      break;
    case "u":
      text.styles.current = styles;
      text.children[0]["underline"] = true;
      break;
    case "s":
      text.children[0]["strike"] = true;
      break;

    default:
      text.__v = 1;
  }
  text.state = currentState;
  return text;
};

const generate = (x) => {
  let children = null;
  if (x.children) {
    children = x.children
      .map((item) => {
        return generate(item);
      })
      .filter((i) => i !== null);
  }

  if (x.type === "root") {
    return {
      section: children,
    };
  } else if (x.type === "p") {
    console.log("file: doc.service.js | line 190 | generate | x", x);
    const test = Object.entries(x.styles)[0];
    let style = {};
    if (test !== null) {
      style = test[1];
    }

    console.log(
      "file: doc.service.js | line 199 | generate | children",
      children
    );

    let newChildren = [];
    children.forEach((c) => {
      if (Array.isArray(c)) {
        newChildren = [...newChildren, ...c];
      } else {
        newChildren.push(c);
      }
    });
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
    return new TextRun({
      text: text.value,
      ...current,
    });
  } else if (x.type === "span") {
    return children;
    // console.log("file: doc.service.js | line 159 | generate | x", x);
  } else if (x.type === "strong") {
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
    return new TextRun({
      text: "COMMENT HERE",
    });
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
    });
  }
  console.log("here", x);
  return null;
};
const createNew = (json) => {
  // const { children } = json;
  const x = traverse(json);
  console.log("file: doc.service.js | line 186 | createNew | x", x);
  let generated = generate(x).section.filter((i) => i);
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            pageNumbers: {
              start: 1,
              formatType: NumberFormat.DECIMAL,
            },
            margin: {
              top: 0,
              right: 200,
              bottom: 0,
              left: 200,
            },
          },
        },
        children: generated,
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
        // saveAs(docblob, "test.docx");
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
