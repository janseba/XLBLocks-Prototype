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
    this.setColour(20);
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
    this.setColour(20);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['range'] = {
  init: function() {
    this.appendDummyInput()
        .appendField(new Blockly.FieldTextInput("range"), "NAME");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['sum'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("SUM");
    this.appendValueInput("NAME")
        .setCheck(null);
    this.setInputsInline(true);
    this.setPreviousStatement(true, null);
    this.setNextStatement(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['for_each_row'] = {
  init: function() {
    this.appendValueInput("NAME")
        .setCheck(null)
        .appendField("EACH ROW IN RANGE");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};