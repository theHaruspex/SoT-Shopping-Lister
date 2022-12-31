const shopping_list_sheet = SpreadsheetApp.getActive().getSheetByName("Shopping List")
const config_sheet = SpreadsheetApp.getActive().getSheetByName("Config")
const inventory_sheet = SpreadsheetApp.getActive().getSheetByName("Inventory")
const section_header_background_hex = "#fbbc04"
const light_yellow_hex = "#fff2cc"
const main_inventory_sheet = SpreadsheetApp.openById("1f0NIkMzBC-7msqJ0g_u3bo8t5g7ht1BifSWRRcWAnh0").getSheetByName("Inventory")
const progress_cell = inventory_sheet.getRange("A13")



function deconstruct_cell(target_sheet, cell_coordinates) {
  let cell = target_sheet.getRange(cell_coordinates[0], cell_coordinates[1])
  let cell_value = cell.getValue()
  let cell_background_hex = cell.getBackground()
  return [cell_value, cell_background_hex]
}


function get_section_row_bounds(target_section_title, target_sheet) {
  let sheet_height = target_sheet.getLastRow() + 1
  let [beginning_row_index, ending_row_index] = [null, null]
  let inside_range = false
  for (let i = 1; i < sheet_height; i++) {
    let [cell_value, cell_background_hex] = deconstruct_cell(target_sheet, [i, 1])
    if (cell_value == target_section_title) {
      beginning_row_index = i
      inside_range = true
    }
    else if (inside_range && cell_background_hex == section_header_background_hex){
      ending_row_index = i - 1
      break
    }
 }
 return [beginning_row_index, ending_row_index]
}


function get_section_range(target_section_title, target_sheet, target_width=null) {
  let [beginning_row, ending_row] = get_section_row_bounds(target_section_title, target_sheet)
  let section_height = ending_row - (beginning_row + 1) // The 1 represents the height of the section header cell.
  let section_width = target_width
  if (section_width == null) {
    section_width = target_sheet.getLastColumn()
  }
  let section_range = target_sheet.getRange(beginning_row, 1, section_height, section_width)
  return section_range
}


function filter_inventory_section_values(inventory_section_values) {
  let filtered_section_values = inventory_section_values.filter(function (inventory_section_values) {
    let item_name = inventory_section_values[0]
    let item_quantity = inventory_section_values[1]
    let clause_1 = item_name != ''
    let clause_2 = typeof(item_quantity) == 'number'
    return (clause_1 && clause_2) 
  })
  return filtered_section_values
}


function get_inventory_section_values(target_section_title) {
  let section_range = get_section_range(target_section_title, inventory_sheet, 2)
  let section_values = section_range.getValues()
  let filtered_section_values = filter_inventory_section_values(section_values)
  return filtered_section_values
}


function get_inventory_section_titles() {
  let first_column_values = inventory_sheet.getRange('A1:A').getValues()
  let section_titles = []
  for (let i = 1; i < first_column_values.length; i++) {
    let [cell_value, cell_background_hex] = deconstruct_cell(inventory_sheet, [i, 1])
    if (cell_background_hex == section_header_background_hex) {
      section_titles.push(cell_value)
    }
  }
  return section_titles
}


function get_other_inventory_values() {
  let other_inventory_section_titles = get_inventory_section_titles().slice(2)
  let other_inventory_values = []
  for (let i = 0; i < other_inventory_section_titles.length; i++) {
    let section_title = other_inventory_section_titles[i]
    let section_values = get_inventory_section_values(section_title)
    other_inventory_values.push(section_values)
  }
  return other_inventory_values.flat(1)
}


function get_complete_inventory_values () {
  let tub_inventory_values = get_inventory_section_values("Ice Cream Tubs")
  let pint_inventory_values = get_inventory_section_values("Pre-Packed Pints")
  let other_inventory_values = get_other_inventory_values()
  return [tub_inventory_values, pint_inventory_values, other_inventory_values]
}


function generate_ideal_inventory_section(target_section_scope) {
  let ideal_inventory_section = {}
  let section_range = config_sheet.getRange(target_section_scope)
  let section_values = section_range.getValues()
  for (let i = 0; i < section_values.length; i++) {
    let [item_name, item_ideal_quantity, item_replace_threshold] = section_values[i]
    if (typeof(item_ideal_quantity) != 'number') {
      continue
    }
    ideal_inventory_section[item_name] = {
      ideal_quantity: item_ideal_quantity,
      replace_threshold: item_replace_threshold
    }
  }
  return ideal_inventory_section
}


function generate_complete_ideal_inventory () {
  let section_scopes_array = ['A3:C', 'E3:G', 'I3:K']
  let ideal_inventory_sections_array = []
  for (let i = 0; i < section_scopes_array.length; i++) {
    let section_scope = section_scopes_array[i]
    let ideal_inventory_section =  generate_ideal_inventory_section(section_scope)
    ideal_inventory_sections_array.push(ideal_inventory_section)
  }
  return ideal_inventory_sections_array
}


function extract_ideal_quantity(ideal_item_entry) {
  try {
    let ideal_quantity = ideal_item_entry['ideal_quantity']
    return ideal_quantity
  }
  catch (TypeError) {
    script_encountered_mismatch_error()
    throw TypeError
  }
}


function calculate_item_deficit(item_values, ideal_inventory_category) {
  let [item_name, item_quantity] = item_values
  let ideal_item_entry = ideal_inventory_category[item_name]
  let ideal_quantity = extract_ideal_quantity(ideal_item_entry)

  let replace_threshold = ideal_item_entry['replace_threshold']
  if (item_quantity > replace_threshold) {
    return
  }
  return (ideal_quantity - item_quantity)
}


function calculate_shopping_list_by_category(inventory_values, ideal_inventory_catagory, category_name) {
  shopping_list = {
    category: category_name,
    needed_items: []
  }
  for (let i = 0; i < inventory_values.length; i++) {
    let item_values = inventory_values[i]
    let item_name = item_values[0]
    let item_deficit = calculate_item_deficit(item_values, ideal_inventory_catagory)
    if (item_deficit == null) {
      continue
    }
    shopping_list['needed_items'].push([item_name, item_deficit])
  }
  return shopping_list
}

function calculate_all_shopping_list_categories() {
  let [tubs_inventory_values, pint_inventory_values, other_inventory_values] = get_complete_inventory_values()
  let [ideal_tubs_inventory, ideal_pints_inventory, ideal_other_inventory] = generate_complete_ideal_inventory()
  let tubs_shopping_list = calculate_shopping_list_by_category(
    tubs_inventory_values,
    ideal_tubs_inventory,
    "Ice Cream Tubs"
  )
  let pint_shopping_list = calculate_shopping_list_by_category(
    pint_inventory_values,
    ideal_pints_inventory,
    "Pre-Packed Pints"
  )
  let other_shopping_list = calculate_shopping_list_by_category(
    other_inventory_values,
    ideal_other_inventory,
    "Other"
  )
  return [tubs_shopping_list, pint_shopping_list, other_shopping_list]
}


function format_section_header(range) {
    range.setFontFamily('Impact')
    range.setFontSize('24')
    range.setBackground(section_header_background_hex)
}


function format_section_entry(range) {
    range.setFontFamily('Comfortaa')
    range.setFontSize('14')
}


function format_last_shopping_list_row() {
  let last_row_index = shopping_list_sheet.getLastRow()
  let last_row = shopping_list_sheet.getRange(last_row_index, 1, 1, 2)
  let column_b_value = last_row.getValues()[0][1]
  if (column_b_value == "Amount Needed") {
    format_section_header(last_row)
  }
  else {
    format_section_entry(last_row)
  }
}


function populate_shopping_list_section(shopping_list) {
  let category = shopping_list['category']
  let needed_items = shopping_list['needed_items']
  shopping_list_sheet.appendRow([category, "Amount Needed"])
  format_last_shopping_list_row()
  for (let i = 0; i < needed_items.length; i++) {
    let entry = needed_items[i]
    shopping_list_sheet.appendRow(entry)
    format_last_shopping_list_row()
  }
}


function populate_shopping_list_sheet() {
  let shopping_lists_array = calculate_all_shopping_list_categories()
  shopping_list_sheet.clear()
  for (let i = 0; i < shopping_lists_array.length; i++) {
    let shopping_list = shopping_lists_array[i]
    populate_shopping_list_section(shopping_list)
  }
}


function transfer_tubs_inventory() {
  let tub_inventory_values = get_inventory_section_values('Ice Cream Tubs')
}


function script_is_processing() {
  progress_cell.setBackground("#d9d9d9")
  progress_cell.setFontColor('black')
  progress_cell.setValue("The script is currently executing. This should take about 30 seconds...")
}

function script_is_complete() {
  progress_cell.setBackground('green')
  progress_cell.setFontColor('white')
  progress_cell.setValue("The script is finished! Have a good night.")
}


function script_encountered_mismatch_error() {
  progress_cell.setBackground("red")
  progress_cell.setFontColor('white')
  progress_cell.setValue("The script encountered a mismatch error! Check that all the inventory items are represented in the config sheet. Also, that they are spelled correctly!")
}


function finish_eod_report() {
  script_is_processing()
  populate_shopping_list_sheet()
  script_is_complete()
}
