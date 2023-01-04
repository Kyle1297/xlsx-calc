"use strict";

module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];

    self.row = function() {
        const args = self.args.map(f => f.args[0]);
        if (args.length > 1) {
            return new Error('#N/A');
        }
        if (args.length === 0 || args[0] == undefined) {
            return parseInt(self.args[0].formula.name.replace(/[^0-9]/g, ''));
        }
    
        if (args[0].name !== 'RefValue' && args[0].name !== 'Range') {
            return new Error('#VALUE!');
        }
    
        if (args[0].name === 'RefValue') {
            const { row_num } = args[0].parseRef();
            return row_num;
        }
    
        if (args[0].name === 'Range') {
            const { min_row, max_row } = args[0].parseRange();
    
            const result = []
            for (let i = min_row; i < max_row + 1; i++) {
                result.push([i]);
            }
    
            return result;
        }
    
        return new Error('#N/A');
    }

    self.calc = function() {
        var errorValues = {
            '#NULL!': 0x00,
            '#DIV/0!': 0x07,
            '#VALUE!': 0x0F,
            '#REF!': 0x17,
            '#NAME?': 0x1D,
            '#NUM!': 0x24,
            '#N/A': 0x2A,
            '#GETTING_DATA': 0x2B
        }, result;
        try {
            if (user_function.name === 'ROW') {
                result = self.row();
            } else {
                result = user_function.apply(self, self.args.map(f => f.calc(user_function.name === "FILTER")));
            }

            if ((user_function.name === "FILTER" || user_function.name === "ROW") && Array.isArray(result) && result.length > 0 && Array.isArray(result[0]) && result[0].length > 0) {
                result = result[0][0];
            }
        } catch (e) {
            if (user_function.name === 'is_blank'
                && errorValues[e.message] !== undefined) {
                // is_blank applied to an error cell doesn't propagate the error
                result = 0;
            }
            else if (user_function.name === 'iserror'
                && errorValues[e.message] !== undefined) {
                // iserror applied to an error doesn't propagate the error and returns true
                result = true;
            }
            else {
                throw e;
            }
        }
        return result;
    };
    self.push = function(buffer) {
        self.args.push(buffer);
    };
};