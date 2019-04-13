Blockly.Blocks['formula'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("Formula")
        .appendField(new Blockly.FieldTextInput("formula name"), "formula_name");
    this.appendValueInput("output")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("formula output");
    this.appendValueInput("statements")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_RIGHT)
        .appendField("functions");
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
    this.setOutput(true, null);
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

Blockly.Blocks['for_each_column'] = {
  init: function() {
    this.appendValueInput("range_each_column_in_range")
        .setCheck(null)
        .appendField("EACH COLUMN IN RANGE");
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(65);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['lookup'] = {
  init: function() {
    this.appendDummyInput()
        .appendField("LOOKUP");
    this.appendValueInput("lookupValue")
        .setCheck(null)
        .appendField("lookup value");
    this.appendValueInput("lookupColumn")
        .setCheck(null)
        .appendField("lookup column");
    this.appendValueInput("resultColumn")
        .setCheck(null)
        .appendField("result column");
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("This functions returns a value from the result column at the same row it finds a match in the lookup column");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['subtract'] = {
  init: function() {
    this.appendValueInput("left_operand")
        .setCheck(null);
    this.appendDummyInput()
        .appendField("-");
    this.appendValueInput("right_operand")
        .setCheck(null);
    this.setInputsInline(false);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};

Blockly.Blocks['divide'] = {
  init: function() {
    this.appendValueInput("numerator")
        .setCheck(null)
        .setAlign(Blockly.ALIGN_CENTRE);
    this.appendDummyInput()
        .appendField("/");
    this.appendValueInput("denominator")
        .setCheck(null);
    this.setInputsInline(true);
    this.setOutput(true, null);
    this.setColour(120);
 this.setTooltip("");
 this.setHelpUrl("");
  }
};
