
import org.benjamin.matchmapgenerator.MatchMapGenerator;

import java.io.File;

/**
 * @Author: Benjamin Zheng
 * @Date: 2020/11/23 18:15
 */
public class GuiYuanBadminton {
    public static void main(String[] args) {
        MatchMapGenerator matchMapGenerator = new MatchMapGenerator();
        matchMapGenerator.setFile(new File("src/test/resources/报名表.txt"));
        matchMapGenerator.setScoringPaper(new File("src/test/resources/男单计分表.xls"));
        matchMapGenerator.setCategory("男单");
        matchMapGenerator.setGroupMemberNumber(10);
        matchMapGenerator.generator();

        matchMapGenerator.setScoringPaper(new File("src/test/resources/女单计分表.xls"));
        matchMapGenerator.setCategory("女单");
        matchMapGenerator.setGroupMemberNumber(5);
        matchMapGenerator.generator();

        matchMapGenerator.setScoringPaper(new File("src/test/resources/健身运动组计分表.xls"));
        matchMapGenerator.setCategory("健身");
        matchMapGenerator.setGroupMemberNumber(4);
        matchMapGenerator.generator();

    }
}
