Blockly.Blocks['rowsum'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("RowSum");
    this.appendValueInput("NAME")
        .setCheck(null)
        .appendField("data block");
    this.setPreviousStatement(true, null);
    this.setColour(230);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['datablock'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("Absolute")
        .appendField(new Blockly.FieldCheckbox("FALSE"), "NAME");
    this.appendDummyInput()
        .appendField("Range")
        .appendField(new Blockly.FieldTextInput("A1:B2"), "RANGE");
    this.setOutput(true, null);
    this.setColour(105);
 this.setTooltip("Enter reference to range in A1 notation");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['columnsum'] = {
  init: function() {
    this.appendDummyInput()
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("ColumnSum");
    this.appendValueInput("DATABLOCK")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("data block");
    this.setPreviousStatement(true, null);
    this.setColour(230);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['formula'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Formula")
        .appendField(new Blockly.FieldTextInput("formula name"), "formula_name");
    this.appendValueInput("NAME")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("formula output");
    this.appendStatementInput("NAME")
        .setCheck(null);
    this.setColour(210);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['definenamedranges'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Define Named Ranges");
    this.appendStatementInput("namedRangeDefinition")
        .setCheck(null);
    this.setColour(210);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['find'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Find");
    this.appendValueInput("NAME")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("what");
    this.appendValueInput("NAME")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("look in");
    this.appendValueInput("NAME")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("return column");
    this.setPreviousStatement(true, null);
    this.setColour(230);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['range'] = {
  init: function() {
    this.appendDummyInput()
        .appendField(new Blockly.FieldTextInput("A1:B2"), "NAME");
    this.setOutput(true, null);
    this.setColour(105);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['sum'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Sum");
    this.appendValueInput("NAME")
        .setCheck(null)
        .appendField("range");
    this.setPreviousStatement(true, null);
    this.setColour(230);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};