package com.github.kuxuxun.scotchtest.assertion.assertion;

import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.scotch.excel.area.ScPos;
import com.github.kuxuxun.scotch.excel.area.ScRange;

public class In {
	String start, end;

	public static In fromScRange(ScRange range) {
		return new In(range.getTopLeft().toReference(),

		range.getBottomRight().toReference());
	}

	public ScRange toScRange() {
		return new ScRange(new ScPos(start), new ScPos(end));
	}

	public In(String start, String end) {
		this.start = start;
		this.end = end;

		// TODO emptyを許すのかどうか再考
		if (n2b(start).length() == 0 || n2b(end).length() == 0) {
			throw new IllegalStateException("ポジションが指定されていませんです start:[" + start
					+ "] end: [" + end + "]");
		}
		toScRange().validRangeOrThrowException();

	}

	private static String n2b(String s) {
		if (s == null)
			return "";
		return s;
	}

	public String start() {
		return start;
	}

	public String end() {
		return end;
	}

	public boolean isLeftEdge(ScPos pos) {
		if (new ScPos(start).getCol() == pos.getCol()) {
			return true;
		}
		return false;
	}

	public boolean isTopEdge(ScPos pos) {
		if (new ScPos(start).getRow() == pos.getRow()) {
			return true;
		}
		return false;
	}

	public boolean isBottomEdge(ScPos pos) {
		if (new ScPos(end).getRow() == pos.getRow()) {
			return true;
		}
		return false;
	}

	public boolean isRightEdge(ScPos pos) {
		if (new ScPos(end).getCol() == pos.getCol()) {
			return true;
		}
		return false;
	}

	public List<ScPos> getPositions() {

		List<ScPos> r = new ArrayList<ScPos>();

		ScPos topLeftCell = new ScPos(start());
		ScPos bottomRightCell = new ScPos(end());

		for (int row = topLeftCell.getRow(); row <= bottomRightCell.getRow(); row++) {
			for (int col = topLeftCell.getCol(); col <= bottomRightCell
					.getCol(); col++) {

				r.add(new ScPos(row, col));

			}
		}

		return r;

	}
}