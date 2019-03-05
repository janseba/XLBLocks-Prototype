Blockly.Blocks['formula'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Formula")
        .appendField(new Blockly.FieldTextInput("formula name"), "formula_name");
    this.appendValueInput("output")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("formula output");
    this.appendStatementInput("statements")
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
        .appendField(new Blockly.FieldTextInput("range"), "range_address");
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
    this.appendValueInput("sum_parameters")
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
    this.appendValueInput("range_each_row_in_range")
        .setCheck(null)
        .appendField("EACH ROW IN RANGE");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};