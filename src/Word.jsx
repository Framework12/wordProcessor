import React, { useState, useEffect, useRef } from "react";
import { saveAs } from "file-saver";
import { Packer, Document, Paragraph } from "docx";

const DEFAULT_FONT_SIZE = 16;
const DEFAULT_FONT_COLOR = "#000000";
const DEFAULT_BACKGROUND_COLOR = "#ffffff";
const DEFAULT_FONT_STYLE = "Arial";

const Word = () => {
  const [content, setContent] = useState("");
  const [fontSize, setFontSize] = useState(DEFAULT_FONT_SIZE);
  const [fontColor, setFontColor] = useState(DEFAULT_FONT_COLOR);
  const [backgroundColor, setBackgroundColor] = useState(
    DEFAULT_BACKGROUND_COLOR
  );
  const [fontStyle, setFontStyle] = useState(DEFAULT_FONT_STYLE);
  const [undoStack, setUndoStack] = useState([]);
  const [redoStack, setRedoStack] = useState([]);
  const textAreaRef = useRef(null);

  useEffect(() => {
    textAreaRef.current.focus();
  }, []);

  useEffect(() => {
    if (content !== "") {
      setUndoStack((prevUndoStack) => [...prevUndoStack, content]);
    }
  }, [content, setUndoStack]);

  const handleChange = (e) => {
    setContent(e.target.value);
  };

  const handleFontSizeChange = (e) => {
    setFontSize(parseInt(e.target.value));
  };

  const handleFontColorChange = (e) => {
    setFontColor(e.target.value);
  };

  const handleBackgroundColorChange = (e) => {
    setBackgroundColor(e.target.value);
  };

  const handleFontStyleChange = (e) => {
    setFontStyle(e.target.value);
  };

  const handleUndo = () => {
    if (undoStack.length > 1) {
      const prevState = undoStack.slice(0, -1);
      const lastState = prevState[prevState.length - 1];
      setRedoStack([content, ...redoStack]);
      setUndoStack(prevState);
      setContent(lastState);
    }
  };

  const handleRedo = () => {
    if (redoStack.length > 0) {
      const nextState = redoStack[0];
      const newRedoStack = redoStack.slice(1);
      setUndoStack([...undoStack, content]);
      setRedoStack(newRedoStack);
      setContent(nextState);
    }
  };

  const handleEraseText = () => {
    setContent("");
  };

  const handleToUpperCase = () => {
    setContent(content.toUpperCase());
  };

  const handleToLowerCase = () => {
    setContent(content.toLowerCase());
  };

  const handleSave = () => {
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [new Paragraph(content)],
        },
      ],
    });

    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, "word.docx");
    });
  };

  const wordCount = content.trim().split(/\s+/).length;

  return (
    <div className="word-container">
      <div className="word-toolbar">
        <label className="word-label" htmlFor="fontSize">
          Font Size:
        </label>
        <input
          id="fontSize"
          type="number"
          className="word-input"
          value={fontSize}
          onChange={handleFontSizeChange}
        />
        <label className="word-label" htmlFor="fontColor">
          Font Color:
        </label>
        <input
          id="fontColor"
          type="color"
          className="word-input"
          value={fontColor}
          onChange={handleFontColorChange}
        />
        <label className="word-label" htmlFor="backgroundColor">
          Background Color:
        </label>
        <input
          id="backgroundColor"
          type="color"
          className="word-input"
          value={backgroundColor}
          onChange={handleBackgroundColorChange}
        />
        <label className="word-label" htmlFor="fontStyle">
          Font Style:
        </label>
        <select
          id="fontStyle"
          className="word-select"
          value={fontStyle}
          onChange={handleFontStyleChange}
        >
          {[
            "Arial",
            "Helvetica",
            "Times New Roman",
            "Courier New",
            "Verdana",
            "Georgia",
            "Tahoma",
            "Trebuchet MS",
            "Impact",
            "Comic Sans MS",
            "Arial Black",
            "Arial Narrow",
            "Lucida Console",
            "Lucida Sans Unicode",
            "Palatino Linotype",
            "Garamond",
            "Book Antiqua",
            "Copperplate",
            "Franklin Gothic Medium",
            "Century Gothic",
            "Cambria",
            "Rockwell",
            "Segoe UI",
            "Optima",
            "Geneva",
            "MS Sans Serif",
            "MS Serif",
            "Palatino",
            "Symbol",
            "Roboto",
            "Open Sans",
            "Lato",
            "Montserrat",
            "Raleway",
            "Source Sans Pro",
            "Ubuntu",
            "Oswald",
            "PT Sans",
            "Noto Sans",
          ].map((fontName) => (
            <option key={fontName} value={fontName}>
              {fontName}
            </option>
          ))}
        </select>

        <button
          className="word-button"
          onClick={handleUndo}
          disabled={undoStack.length <= 1}
        >
          Undo
        </button>
        <button
          className="word-button"
          onClick={handleRedo}
          disabled={redoStack.length === 0}
        >
          Redo
        </button>
        <button className="word-button" onClick={handleEraseText}>
          Erase Text
        </button>
        <button className="word-button" onClick={handleToUpperCase}>
          Uppercase
        </button>
        <button className="word-button" onClick={handleToLowerCase}>
          Lowercase
        </button>
        <button className="word-button" onClick={handleSave}>
          Save
        </button>
      </div>
      <textarea
        ref={textAreaRef}
        value={content}
        onChange={handleChange}
        className="word-textarea"
        style={{
          fontSize: `${fontSize}px`,
          color: fontColor,
          backgroundColor,
          fontFamily: fontStyle,
        }}
        placeholder="Start typing here..."
      />

      <div className="word-count">Word Count: {wordCount}</div>
    </div>
  );
};

export default Word;
