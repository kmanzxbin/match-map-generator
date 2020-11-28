
import org.benjamin.matchmapgenerator.MatchMapGenerator;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: Benjamin Zheng
 * @Date: 2020/11/23 18:15
 */
public class GuiYuanBadminton {
    public static void main(String[] args) {
        MatchMapGenerator matchMapGenerator = new MatchMapGenerator();
        matchMapGenerator.setFile(new File("src/test/resources/报名表.txt"));

        generator(matchMapGenerator, "男单", 5);
        generator(matchMapGenerator, "女单", 5);
        generator(matchMapGenerator, "健身休闲", 5);

        matchMapGenerator.setCategory(null);
        List<String> list = matchMapGenerator.getList("男双");
        generator(matchMapGenerator, 4, list.size() / 2 + (list.size() % 2 ==0? 0 : 1), "男双");
        list = matchMapGenerator.getList("女双");
        generator(matchMapGenerator, 4, list.size() / 2 + (list.size() % 2 ==0? 0 : 1), "女双");

        // todo 增加输出分组名单的功能
    }

    static void generator(MatchMapGenerator matchMapGenerator, String category, int groupMemberNumber) {
        matchMapGenerator.setScoringPaper(new File(String.format("src/test/resources/%s分组计分表.xls", category)));
        matchMapGenerator.setCategory(category);
        matchMapGenerator.setGroupMemberNumber(groupMemberNumber);
        matchMapGenerator.generator();
    }

    static void generator(MatchMapGenerator matchMapGenerator, int groupMemberNumber, int memberNumber, String prefix) {

        List<String> list = new ArrayList();
        for (int i = 0; i < memberNumber; i++) {
            list.add(prefix + (i+1));
        }
        matchMapGenerator.setScoringPaper(new File(String.format("src/test/resources/%s分组计分表.xls", prefix)));
        matchMapGenerator.setCategory(prefix);
        matchMapGenerator.setGroupMemberNumber(groupMemberNumber);
        matchMapGenerator.generator(list);
    }
}
