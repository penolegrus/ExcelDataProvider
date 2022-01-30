import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;
import org.testng.Assert;
import org.testng.annotations.Test;

import static io.restassured.RestAssured.given;

public class ExcelTest {

    //чтение эксель файла с помощью массива аргументов
    @Test(dataProvider = "usersFromSheet1", dataProviderClass = ExcelDataProviders.class)
    public void readExcelSheetDefault(String... params){
        System.out.println("Id: " + params[0] + " Имя: " + params[1] + " Фамилия: " + params[2]);
    }

    //чтение эксель файла с кастомного листа с использованием точных аргументов
    @Test(dataProvider = "usersFromSheet2", dataProviderClass = ExcelDataProviders.class)
    public void readExcelSheet2(String params1, String params2){
        System.out.println("Логин:" + params1);
        System.out.println("Пароль:" + params2);
    }

    //тест на проверку пользователей через Response
    @Test(dataProvider = "usersForApi", dataProviderClass = ExcelDataProviders.class)
    public void checkUsers(String... params){
        int id = (int) Double.parseDouble(params[0]);
        Response response = given()
                .contentType(ContentType.JSON)
                .get("https://reqres.in/api/users/"+id)
                .then().log().body().extract().response();
        JsonPath jsonPath = response.jsonPath();
        Assert.assertEquals(jsonPath.getInt("data.id"),id);
        Assert.assertEquals(jsonPath.getString("data.email"),params[1]);
        Assert.assertEquals(jsonPath.getString("data.first_name"),params[2]);
        Assert.assertEquals(jsonPath.getString("data.last_name"),params[3]);
        Assert.assertEquals(jsonPath.getString("data.avatar"),params[4]);
    }

    //тест на проверку пользователей через Pojo
    @Test(dataProvider = "usersForApi", dataProviderClass = ExcelDataProviders.class)
    public void checkUsersPojo(String... params){
        int id = (int) Double.parseDouble(params[0]);
        UserPojo user = given()
                .contentType(ContentType.JSON)
                .get("https://reqres.in/api/users/"+id)
                .then().log().body()
                .extract().body().jsonPath().getObject("data",UserPojo.class);
        Assert.assertEquals(user.getId(),id);
        Assert.assertEquals(user.getEmail(),params[1]);
        Assert.assertEquals(user.getFirst_name(),params[2]);
        Assert.assertEquals(user.getLast_name(),params[3]);
        Assert.assertEquals(user.getAvatar(),params[4]);
    }

}
