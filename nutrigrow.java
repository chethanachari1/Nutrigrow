import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelSearchApp {

public static void main(String[] args) {
SwingUtilities.invokeLater(() -> {
GuiFrame gui = new GuiFrame();
gui.setVisible(true);
});
}

//Finding out level of malnutrition
public static void searchExcel(double targetLength, double targetValue) {
// path of Excel file
String excelFilePath = "D://Introproj//data.xlsx";
// Initialize matchedColumn to null
String matchedColumn = null;

try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
Workbook workbook = WorkbookFactory.create(fileInputStream)) {
Sheet sheet = workbook.getSheetAt(0);
int lengthColumnIndex = -1;
Row headerRow = sheet.getRow(0);

for (Cell cell : headerRow) {
String header = cell.getStringCellValue();
if (header.equalsIgnoreCase("length")) {
lengthColumnIndex = cell.getColumnIndex();
}
}

if (lengthColumnIndex == -1) {
throw new IllegalArgumentException("Missing required 'length' column in the Excel sheet.");
}
// Larger epsilon value for double comparison
double epsilon = 1e-6;

for (Row row : sheet) {
if (row.getRowNum() == 0) {
// Skip the header row
continue;
}

Cell lengthCell = row.getCell(lengthColumnIndex);
if (lengthCell.getCellType() == CellType.NUMERIC) {
double lengthValue = lengthCell.getNumericCellValue();
if (Math.abs(lengthValue - targetLength) < epsilon) {
for (Cell cell : row) {
if (cell.getColumnIndex() != lengthColumnIndex) {
if (isInRange(cell, targetValue)) {
matchedColumn = headerRow.getCell(cell.getColumnIndex()).getStringCellValue();
// Stops processing once a match is found
break;
}
}
}
// Stops processing once a match is found
break;
}
}
}

if (matchedColumn != null) {
String resultMessage = getResultMessage(matchedColumn);
showMessageDialog("Search Result", resultMessage);

String ageString = JOptionPane.showInputDialog(null, "Enter the age in months:");
if (ageString != null && !ageString.isEmpty()) {
int ageMonths = Integer.parseInt(ageString);
showDietPlan(matchedColumn, ageMonths);
}
} else {
showMessageDialog("Search Result", "No matching entry found for the specified length and value.");
}

} catch (IOException e) {
e.printStackTrace();
}
}

private static boolean isInRange(Cell cell, double targetValue) {
String cellValue = cell.getStringCellValue();
String[] rangeParts = cellValue.split("-");
double minRange = Double.parseDouble(rangeParts[0]);
double maxRange = Double.parseDouble(rangeParts[1]);
return targetValue >= minRange && targetValue <= maxRange;
}

private static String getResultMessage(String matchedColumn) {
if (matchedColumn.equals("asam") || matchedColumn.equals("sam")) {
return "The child is having severe acute malnutrition";
} else if (matchedColumn.equals("mam")) {
return "The child is having moderate acute malnutrition";
} else if (matchedColumn.equals("sd1")) {
return "The child is having mild malnutrition";
} else {
return "Child is normal";
}
}

private static void showDietPlan(String matchedColumn, int ageMonths) {
StringBuilder dietPlanMessage = new StringBuilder();

switch (matchedColumn) {
case "mam":
dietPlanMessage.append("Moderate Acute Malnutrition Diet Plan:\n");
handleMAM(dietPlanMessage, ageMonths);
break;
case "sd1":
dietPlanMessage.append("Mild Malnutrition Diet Plan:\n");
handleSD1(dietPlanMessage, ageMonths);
break;
case "median":
dietPlanMessage.append("Maintain the regular diet.\n");
break;
default:
dietPlanMessage.append("Consult a pediatrician.\n");
break;
}

showMessageDialog("Diet Plan", dietPlanMessage.toString());
}

//diet plans for different levels of malnutrition based on child's age
private static void handleMAM(StringBuilder dietPlanMessage, int ageMonths) {
if (ageMonths <= 6) {
dietPlanMessage.append("Moderately Acute Malnutrition Diet Plan for children younger than six months:\n");
dietPlanMessage
.append("> Breast feed as often as child wants, day and night, at least 8 times in 24 hours\n");
dietPlanMessage.append(
"> Consult a doctor, as moderate malnutrition for this group of children is due to diseases like T.b, diarrhea, lung infections, and stomach infections. If necessary, provide child with zinc and iron supplements.\n");
dietPlanMessage.append("Add the complete diet plan details for age <= 6 months here.\n");
} else if (ageMonths > 6 && ageMonths <= 12) {
dietPlanMessage
.append("Moderately Acute Malnutrition Diet Plan for children between six and twelve months:\n");
dietPlanMessage.append("> Breast milk: on demand feeding approximately 4-6 times a day\n");
dietPlanMessage.append(
"> Complementary Foods: start with few spoonfuls (1 or 2 tablespoons) of mashed fruits, vegetables, or cereals once or twice a day\n");
dietPlanMessage.append("> Gradually increase portion size as the baby's appetite grows\n");
dietPlanMessage.append("Add the complete diet plan details for age between 6 and 12 months here.\n");
} else if (ageMonths > 12 && ageMonths <= 24) {
dietPlanMessage.append(
"Moderately Acute Malnutrition Diet Plan for children between twelve and twenty-four months:\n");
dietPlanMessage.append(
"> Grains: 1/4th to 1/3rd cup of cooked grains (rice, millets, ragi, ragi malt) per meal\n");
dietPlanMessage.append(
"> Protein Source: 1-2 tablespoons of cooked lentils or beans, or 1/2 an egg, or 1-2 ounces of cooked meat or fish\n");
dietPlanMessage.append("> Vegetables: 1/4 to 1/3 cup of cooked vegetables per meal\n");
dietPlanMessage.append("> Fruits: 1/4 to 1/3 cup of chopped or sliced fruits per meal\n");
dietPlanMessage.append("> Dairy: 1/2 to 3/4 cup of milk or curd per day\n");
dietPlanMessage.append("Add the complete diet plan details for age between 12 and 24 months here.\n");
} else {
dietPlanMessage
.append("Moderately Acute Malnutrition Diet Plan for children older than twenty-four months:\n");
dietPlanMessage.append("> Grains: 1/3 to 1/2 cup of cooked grains per meal\n");
dietPlanMessage.append(
"> Protein Sources: 2-3 tablespoons of cooked lentils or beans, or 1 egg, 2-3 ounces of cooked meat or fish\n");
dietPlanMessage.append("> Vegetables: 1/3 to 1/2 cup of cooked vegetables per meal\n");
dietPlanMessage.append("> Fruits: 1/3 to 1/2 cup of chopped or sliced fruits per meal\n");
dietPlanMessage.append("> Dairy: 1 cup of milk or curd per day\n");
dietPlanMessage.append("Add the complete diet plan details for age older than 24 months here.\n");
}
dietPlanMessage.append(
"Ready-to-use therapeutic foods (RUTFs) are high-energy, high-protein pastes. They are a good option for children above 6 months.\n");
dietPlanMessage.append("Here is a recipe for a RUTF made from millets and peanuts:\n");
dietPlanMessage.append("Ingredients:\n");
dietPlanMessage.append("> 1 cup millet flour\n");
dietPlanMessage.append("> 1/2 cup peanut butter\n");
dietPlanMessage.append("> 1/4 cup sugar\n");
dietPlanMessage.append("> 1/4 cup oil\n");
dietPlanMessage.append("> 1 tsp salt\n");
dietPlanMessage.append("Instructions:\n");
dietPlanMessage.append("1. In a blender, combine all the ingredients and blend until smooth\n");
dietPlanMessage.append("2. Pour the mixture into a bowl and store in the refrigerator\n");
dietPlanMessage.append("NOTE: RUTF can be stored in the refrigerator for up to two weeks.\n");
}

private static void handleSD1(StringBuilder dietPlanMessage, int ageMonths) {
if (ageMonths <= 6) {
dietPlanMessage.append("Mild Malnutrition Diet Plan for children younger than six months:\n");
dietPlanMessage
.append("> Breast feed as often as child wants, day and night, at least 8 times in 24 hours\n");
} else if (ageMonths > 6 && ageMonths <= 12) {
dietPlanMessage.append("Mild Malnutrition Diet Plan for children between six and twelve months:\n");
dietPlanMessage.append("> Breast feed as often as child wants\n");
dietPlanMessage.append("> Give at least one katori serving at a time:\n");
dietPlanMessage.append("  - Mashed Roti/Rice/Bread/Biscuit mixed in sweetened undiluted Milk\n");
dietPlanMessage.append(
"  - Mashed Roti/rice/Bread mixed in Thick dal with added ghee/Oil or Khichdi with added oil or ghee, add cooked vegetables also in the servings\n");
dietPlanMessage
.append("  - Sevian / Dalia/Halwa/Kheer prepared in Milk or any Cereal porridge cooked in Milk\n");
dietPlanMessage.append("  - Mashed boiled/Fried Potato\n");
dietPlanMessage.append(
"> Also give Nutritious Food between meals, such as: Banana/Biscuits/cheeko/Mango/Papaya as snacks\n");
dietPlanMessage.append("> 3 times per day if breast feed, 5 times per day if not breast feed\n");
} else if (ageMonths > 12 && ageMonths <= 24) {
dietPlanMessage.append("Mild Malnutrition Diet Plan for children between twelve and twenty-four months:\n");
dietPlanMessage.append("> Breast feed as often as the child wants\n");
dietPlanMessage.append("> Offer food from the family pot\n");
dietPlanMessage.append("> Give at least one and a half katori servings at a time of:\n");
dietPlanMessage.append(
"  - Mashed roti/rice/bread mixed in thick dal with added ghee/oil or Khichdi with added oil/ghee, add cooked vegetables also in the servings\n");
dietPlanMessage.append("  - Mashed roti/rice/bread/biscuit mixed in sweetened undiluted milk\n");
dietPlanMessage
.append("  - Sevian / Dalia/Halwa/Kheer prepared in Milk or any Cereal porridge cooked in Milk\n");
dietPlanMessage.append("  - Mashed boiled/Fried Potato\n");
dietPlanMessage.append(
"> Also give Nutritious Food between meals, such as: Banana/Biscuits/cheeko/Mango/Papaya as snacks\n");
dietPlanMessage.append("5 times per day\n");
} else {
dietPlanMessage.append("Mild Malnutrition Diet Plan for children older than twenty-four months:\n");
dietPlanMessage.append("> Give family foods at 3 meals each day\n");
dietPlanMessage.append(
"> Also give twice daily Nutritious Food between meals, such as: Banana/Biscuits/cheeko/Mango/Papaya as snacks\n");
}
}

private static void showMessageDialog(String title, String message) {
JOptionPane.showMessageDialog(null, message, title, JOptionPane.INFORMATION_MESSAGE);
}
}

//GUI Implementation
class GuiFrame extends JFrame implements ActionListener {
Container contentPane;
JButton searchButton;
JTextField lengthTextField;
JTextField valueTextField;
JLabel h;
JLabel w;
JPanel panel, panel1;

public GuiFrame() {
super("NutriGrow");

contentPane = this.getContentPane();

panel = new JPanel();
panel1 = new JPanel();
panel1.setLayout(new FlowLayout(FlowLayout.CENTER));
panel.setLayout(new GridBagLayout());
GridBagConstraints gbc = new GridBagConstraints();
h = new JLabel("Enter height in cm");
gbc.gridx = 0;
gbc.gridy = 0;

panel.add(h, gbc);
gbc.weightx = 0.01;
lengthTextField = new JTextField(20);
gbc.gridx = 1;
gbc.gridy = 0;

panel.add(lengthTextField, gbc);
w = new JLabel("Enter weight in kg");
gbc.gridx = 0;
gbc.gridy = 1;
panel.add(w, gbc);
gbc.weightx = 0.01;
valueTextField = new JTextField(20);
gbc.gridx = 1;
gbc.gridy = 1;

panel.add(valueTextField, gbc);

searchButton = new JButton("Search");
gbc.gridx = 1;
gbc.gridy = 2;
panel.add(searchButton, gbc);

contentPane.add(panel, BorderLayout.CENTER);
contentPane.add(panel1, BorderLayout.SOUTH);

searchButton.addActionListener(this);

setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
setSize(400, 300);
setResizable(false);
}

@Override
public void actionPerformed(ActionEvent e) {
double targetLength = Double.parseDouble(lengthTextField.getText());
double targetValue = Double.parseDouble(valueTextField.getText());

ExcelSearchApp.searchExcel(targetLength, targetValue);

}
}