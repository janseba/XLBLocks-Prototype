Blockly.JavaScript['formula'] = function(block) {
  var text_formula_name = block.getFieldValue('formula_name');
  var value_name = Blockly.JavaScript.valueToCode(block, 'NAME', Blockly.JavaScript.ORDER_ATOMIC);
  var statements_name = Blockly.JavaScript.statementToCode(block, 'NAME');
  // TODO: Assemble JavaScript into code variable.
  var code = '...;\n';
  return code;
};

Blockly.JavaScript['definenamedranges'] = function(block) {
  var statements_namedrangedefinition = Blockly.JavaScript.statementToCode(block, 'namedRangeDefinition');
  // TODO: Assemble JavaScript into code variable.
  var code = '...;\n';
  return code;
};

Blockly.JavaScript['range'] = function(block) {
  var text_range_address = block.getFieldValue('range_address');
  // TODO: Assemble JavaScript into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_NONE];
};

Blockly.JavaScript['sum'] = function(block) {
  var value_sum_parameters = Blockly.JavaScript.valueToCode(block, 'sum_parameters', Blockly.JavaScript.ORDER_ATOMIC);
  // TODO: Assemble JavaScript into code variable.
  var code = '...;\n';
  return code;
};

Blockly.JavaScript['for_each_row'] = function(block) {
  var value_range_each_row_in_range = Blockly.JavaScript.valueToCode(block, 'range_each_row_in_range', Blockly.JavaScript.ORDER_ATOMIC);
  // TODO: Assemble JavaScript into code variable.
  var code = '...';
  // TODO: Change ORDER_NONE to the correct strength.
  return [code, Blockly.JavaScript.ORDER_NONE];
};