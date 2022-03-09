<template>
  <div>
    <form @submit.prevent="submitForm">
      <input type="file" @change="fileUploaded" />
      <button type="submit">Start</button>
    </form>
  </div>
</template>

<script>
import { parse } from "@/core/services/htmlParser.service";

let whitelist = [
  "id",
  "tagName",
  "className",
  "childNodes",
  "rawText",
  "attributes",
  "_rawAttrs",
  "rawAttrs",
];
function domToObj(domEl) {
  var obj = {};
  for (let i = 0; i < whitelist.length; i++) {
    if (domEl[whitelist[i]] instanceof NodeList) {
      obj[whitelist[i]] = Array.from(domEl[whitelist[i]]);
    } else {
      obj[whitelist[i]] = domEl[whitelist[i]];
    }
  }
  return obj;
}
// import domJSON from "domjson";
export default {
  name: "TestComponent",
  data() {
    return {
      file: null,
      text: null,
      parsedText: "",
    };
  },
  methods: {
    async submitForm() {
      console.log("Submit Form.");
    },
    fileUploaded(event) {
      const [f] = event.target.files;
      const reader = new FileReader();
      reader.onload = (() => (e) => {
        this.text = e.target.result;
      })(f);
      // Read in the image file as a data URL.
      reader.readAsText(f);
    },
  },
  watch: {
    text(newVal) {
      if (newVal !== null) {
        const parsedText = parse(newVal);
        console.log("file: Test.vue | line 60 | text | parsedText", parsedText);
        // const newText = domJSON.toJSON(parsedText);
        const newText = JSON.stringify(parsedText, function (name, value) {
          if (name === "") {
            return domToObj(value);
          }
          if (Array.isArray(this)) {
            if (typeof value === "object") {
              return domToObj(value);
            }
            return value;
          }
          if (whitelist.find((x) => x === name)) return value;
        });

        console.log(
          "file: Test.vue | line 41 | text | newText",
          JSON.parse(newText).childNodes
        );

        // console.log(
        //   "file: Test.vue | line 39 | text | this.parsedText",
        //   this.parsedText
        // );
      }
    },
  },
};
</script>

<style lang="scss" scoped></style>
