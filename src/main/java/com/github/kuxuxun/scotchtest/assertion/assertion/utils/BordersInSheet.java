package com.github.kuxuxun.scotchtest.assertion.assertion.utils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.github.kuxuxun.scotch.excel.ScSheet;
import com.github.kuxuxun.scotch.excel.area.ScArea;
import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.cell.ScStyle;
import com.github.kuxuxun.scotchtest.assertion.assertion.utils.CellBorders.Border;

public class BordersInSheet {

	public static BordersInSheet fetchFromSheet(String topLeftCell,
			String bottomRightCell, ScSheet sheet) {
		return fetchFromSheet(new ScPos(topLeftCell),
				new ScPos(bottomRightCell), sheet);
	}

	public static BordersInSheet fetchFromSheet(ScPos topLeftCell,
			ScPos bottomRightCell, ScSheet sheet) {
		BordersInSheet result = new BordersInSheet();
		for (int row = topLeftCell.getRow(); row <= bottomRightCell.getRow(); row++) {
			for (int col = topLeftCell.getCol(); col <= bottomRightCell
					.getCol(); col++) {
				ScPos pos = new ScPos(row, col);
				result.getFromCellStyle(pos, sheet.getCellAt(pos).getStyle());
			}
		}

		return result;
	}

	public static enum Destination {
		LEFT, RIGHT, TOP, BOTTOM;
	}

	private final Map<ScPos, CellBorders> cellMatrix = new HashMap<ScPos, CellBorders>();

	public void setBorderdAt(ScPos pos, Destination dest, Border border) {
		CellBorders cell = CellBorders.createOrRegisterdCell(pos, cellMatrix);

		switch (dest) {
		case LEFT:
			cell.setOnLeft(border);
			break;
		case RIGHT:
			cell.setOnRight(border);
			break;

		case TOP:
			cell.setOnTop(border);
			break;

		case BOTTOM:
			cell.setOnBottom(border);
			break;

		default:
			break;
		}

	}

	public Boolean hasBorderedAt(ScPos cellPosition, Destination dest,
			Border simple) {

		CellBorders cell = cellMatrix.get(cellPosition);

		if (cell == null) {
			return false;
		}

		Border border = null;

		switch (dest) {
		case LEFT:
			border = cell.getOnLeft();
			break;
		case RIGHT:
			border = cell.getOnRight();
			break;

		case TOP:
			border = cell.getOnTop();
			break;

		case BOTTOM:
			border = cell.getOnBottom();
			break;

		default:
			break;
		}

		if (border != null && border.equals(simple)) {
			return true;
		}

		return false;

	}

	public void getFromCellStyle(ScPos pos, ScStyle style) {
		CellBorders cell = CellBorders.createOrRegisterdCell(pos, cellMatrix);
		cell.getFromCellStyle(style);
	}

	public Boolean isAroundedIn(ScArea cellArea, Border border) {

		for (int row = cellArea.getTopRow(); row <= cellArea.getBottomRow(); row++) {
			for (int col = cellArea.getLeftCol(); col <= cellArea.getRightCol(); col++) {
				CellBorders cell = CellBorders.createOrRegisterdCell(new ScPos(
						row, col), cellMatrix);

				if (cell == null) {
					return false;
				}

				if (row == cellArea.getTopRow()) {
					if (cell.getOnTop() == null
							|| !cell.getOnTop().equals(border)) {
						return false;
					}
				} else if (row == cellArea.getBottomRow()) {
					if (cell.getOnBottom() == null
							|| !cell.getOnBottom().equals(border)) {
						return false;
					}
				}

				if (col == cellArea.getLeftCol()) {
					if (cell.getOnLeft() == null
							|| !cell.getOnLeft().equals(border)) {
						return false;
					}
				} else if (col == cellArea.getRightCol()) {
					if (cell.getOnRight() == null
							|| !cell.getOnRight().equals(border)) {
						return false;
					}
				}

			}
		}

		return true;
	}

	public Map<ScPos, CellBorders> getMatrix() {
		return this.cellMatrix;
	}

	public List<String> getLookAndFeelDifferenceWith(BordersInSheet counter) {
		Map<ScPos, CellBorders> cellsOfCounter = new HashMap<ScPos, CellBorders>(
				counter.getMatrix());

		List<String> result = new ArrayList<String>();
		for (ScPos eachKey : this.getMatrix().keySet()) {

			if (!cellsOfCounter.containsKey(eachKey)) {
				result.add(eachKey + " is missing in counter");
				continue;
			}

			if (this.cellMatrix.get(eachKey)
					.equals(cellsOfCounter.get(eachKey))) {
				cellsOfCounter.remove(eachKey);
			} else {
				result.add(this.cellMatrix.get(eachKey).getDiffrenceWith(
						cellsOfCounter.get(eachKey)));
			}
		}

		if (cellsOfCounter.size() != 0) {
			for (ScPos each : cellsOfCounter.keySet()) {
				result.add(each + " is missing in expected:");

			}
		}

		return result;

	}

	public Boolean hasEqualLookAndFeelWith(BordersInSheet counter) {
		Map<ScPos, CellBorders> cellsOfCounter = new HashMap<ScPos, CellBorders>(
				counter.getMatrix());

		for (ScPos eachKey : this.getMatrix().keySet()) {
			if (this.cellMatrix.get(eachKey)
					.equals(cellsOfCounter.get(eachKey))) {
				cellsOfCounter.remove(eachKey);
			} else {
				return false;
			}
		}

		if (cellsOfCounter.size() != 0) {
			return false;
		}

		return true;
	}

	@Override
	public String toString() {
		String result = "";
		for (ScPos eachKey : this.getMatrix().keySet()) {
			result += this.cellMatrix.get(eachKey).toString();
		}

		return result;
	}

}
