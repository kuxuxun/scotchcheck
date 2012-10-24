package com.github.kuxuxun.scotchtest.assertion.assertion;

import java.util.ArrayList;
import java.util.List;

import com.github.kuxuxun.scotch.excel.area.ScPos;

//TODO このクラスの意味あるか?
public class Positions extends In {

	private final List<ScPos> positions;

       private Positions(String[] poses) {
		super(poses[0], poses[1]);
		positions = new ArrayList<ScPos>();

		for (String each : poses) {
			positions.add(new ScPos(each));
		}
	}

	public static Positions in(String[] positions) {
	    if(positions.length < 2){
		throw new IllegalStateException("positionsの数は2以上必要です。 but " + positions.length);
	    }
	    return new Positions(positions);
	}

	@Override
	public List<ScPos> getPositions() {
		return this.positions;
	}

}
