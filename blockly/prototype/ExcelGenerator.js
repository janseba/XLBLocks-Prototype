Blockly.JavaScript['rowsum'] = function(block) {
  var variable_data_block = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('data_block'), Blockly.Variables.NAME_TYPE);
  var variable_formula_output = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('formula_output'), Blockly.Variables.NAME_TYPE);
  // TODO: Assemble JavaScript into code variable.
  var code = '...;\n';
  return code;
};

Blockly.JavaScript['datablock'] = function(block) {
  var text_range = block.getFieldValue('RANGE');
  // TODO: Assemble JavaScript into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_NONE];
};

Blockly.JavaScript['columnsum'] = function(block) {
  var value_datablock = Blockly.JavaScript.valueToCode(block, 'DATABLOCK', Blockly.JavaScript.ORDER_ATOMIC);
  var value_formula_output = Blockly.JavaScript.valueToCode(block, 'FORMULA_OUTPUT', Blockly.JavaScript.ORDER_ATOMIC);
  // TODO: Assemble JavaScript into code variable.
  var code = '...;\n';
  return code;
};