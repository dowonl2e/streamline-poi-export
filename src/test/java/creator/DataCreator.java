package creator;

import dto.MemberDto;
import dto.TravelDto;

import java.security.SecureRandom;
import java.time.LocalDateTime;
import java.time.Month;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.concurrent.ThreadLocalRandom;

public class DataCreator {

  private static final String CHARACTERS = "abcdefghijklmnopqrstuvwxyz0123456789";
  private static final SecureRandom RANDOM = new SecureRandom();
  private static final LocalDateTime TODAY = LocalDateTime.now();

  public static List<MemberDto> findMemberAndTravels(int memberCount, int travelCount){
    List<MemberDto> members = new ArrayList<>();
    for(int i = 0 ; i < memberCount ; i++){
      int withdrawValue = RANDOM.nextInt(10);
      MemberDto member = new MemberDto();
      member.setId((long)(i+1));
      member.setName("회원명_"+(i+1));
      member.setEmail(generateString(8)+"@gmail.com");
      member.setTravels(findTravels(travelCount, member.getName()));
      members.add(member);
    }
    return members;
  }

  public static List<TravelDto> findTravels(int travelCount, String name){
    List<TravelDto> travels = new ArrayList<>();
    Random random = new Random();
    for(int i = 0 ; i < travelCount ; i++){
      TravelDto travel = new TravelDto();
      travel.setTravelName(name + "-여행명_"+(i+1));
      int randomValue = random.nextInt(6);
      LocalDateTime now = generateRandomDateTime(2021, Month.JANUARY, 2025, Month.APRIL);
      travel.setStartDate(now);
      travel.setEndDate(now.plusDays(randomValue));
      travel.setCreateDate(now.minusDays(random.nextInt(90)));
      travels.add(travel);
    }
    return travels;
  }

  public static List<MemberDto> findMembers(int memberCount){
    List<MemberDto> members = new ArrayList<>();
    for(int i = 0 ; i < memberCount ; i++){
      int withdrawValue = RANDOM.nextInt(10);
      MemberDto member = new MemberDto();
      member.setId((long)(i+1));
      member.setName("회원명_"+(i+1));
      member.setEmail(generateString(8)+"@gmail.com");
      members.add(member);
    }
    return members;
  }

  public static List<TravelDto> findTravels(int travelCount){
    List<TravelDto> travels = new ArrayList<>();
    Random random = new Random();
    for(int i = 0 ; i < travelCount ; i++){
      TravelDto travel = new TravelDto();
      travel.setTravelName("여행명_"+(i+1));
      int randomValue = random.nextInt(6);
      LocalDateTime now = generateRandomDateTime(2021, Month.JANUARY, 2025, Month.DECEMBER);
      travel.setStartDate(now);
      travel.setEndDate(now.plusDays(randomValue));
      travel.setCreateDate(now.minusDays(random.nextInt(90)));
      travels.add(travel);
    }
    return travels;
  }

  public static LocalDateTime generateRandomDateTime(
      int startYear, Month startMonth, int endYear, Month endMonth
  ) {
    // 시작: 2025-04-01 00:00:00
    LocalDateTime start = LocalDateTime.of(startYear, startMonth, 1, 0, 0);
    // 종료: 2025-06-30 23:59:59
    LocalDateTime end = LocalDateTime.of(endYear, endMonth, 27, 23, 59, 59);

    // 시작~종료 간의 초 단위 차이 계산
    long seconds = ChronoUnit.SECONDS.between(start, end);

    // 랜덤 초 수 생성
    long randomSeconds = ThreadLocalRandom.current().nextLong(seconds + 1);

    // 시작 시간에 랜덤 초를 더해 LocalDateTime 생성
    return start.plusSeconds(randomSeconds);
  }

  private static String generateString(int len) {
    StringBuilder sb = new StringBuilder(len);
    for (int i = 0; i < len; i++) {
      int index = RANDOM.nextInt(CHARACTERS.length());
      sb.append(CHARACTERS.charAt(index));
    }
    return sb.toString();
  }

  private static String generatePhoneNumber() {
    StringBuilder sb = new StringBuilder("010");
    for (int i = 0; i < 8; i++) {
      sb.append(RANDOM.nextInt(10)); // 0~9 중 랜덤 숫자 추가
    }
    return sb.toString(); // 예: 01034892176
  }


}
