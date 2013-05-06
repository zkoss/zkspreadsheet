package zss.testapp.formula;

public class CurrencyFns {
	public static double toTWD(double usdNum, double twdRate) {
        return usdNum * twdRate;
    }
}
