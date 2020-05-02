package cc.cafetime;

public class Main {

	public static void main(String[] args) throws Exception {
		 Img2ByteArray service = new Img2ByteArray();
		 int[][][] result = service.getImagePixel("/Users/steven/Desktop/2.png");
		 System.out.println(result);
		 for (int i = 0; i < result.length; i++) {
			for (int j = 0; j < result[i].length; j++) {
				System.out.print(result[i][j][0] +","+ result[i][j][1] +","+ result[i][j][2] + " \t");
			}
			System.out.println("\n");
		}
		// 写入数据到工作簿对象内
		ExcelWriter.writeData(result, "/Users/steven/Desktop/flag_2.xlsx");
	}
}
