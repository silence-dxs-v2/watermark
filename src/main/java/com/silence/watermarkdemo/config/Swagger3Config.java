package com.silence.watermarkdemo.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import springfox.documentation.builders.ApiInfoBuilder;
import springfox.documentation.builders.PathSelectors;
import springfox.documentation.builders.RequestHandlerSelectors;
import springfox.documentation.builders.RequestParameterBuilder;
import springfox.documentation.oas.annotations.EnableOpenApi;
import springfox.documentation.schema.ScalarType;
import springfox.documentation.service.ApiInfo;
import springfox.documentation.service.ParameterType;
import springfox.documentation.service.RequestParameter;
import springfox.documentation.spi.DocumentationType;
import springfox.documentation.spring.web.plugins.Docket;

import java.util.ArrayList;
import java.util.List;

@Configuration
@EnableOpenApi // swagger3开启注解，与swagger2不同
public class Swagger3Config {

    /**
     * 文档基本信息
     */
    private ApiInfo apiInfo() {
        return new ApiInfoBuilder()
                .title("swagger口文档")
                .description("swagger接口文档")
                .version("2.0")
                .build();
    }

    /**
     * 全局通用属性（摘要）配置
     */
    @Bean
    public Docket api() {
        return new Docket(DocumentationType.OAS_30)
                .apiInfo(apiInfo()) //应用文档基本信息
                .select()
                .apis(RequestHandlerSelectors.basePackage("com.silence.watermarkdemo.controller")) // swagger扫描路径
                .paths(PathSelectors.any()) // 应用于包下所有路径
                .build();
//                .ignoredParameterTypes(User.class, AdminUser.class) // 忽略此类型输入参数（viewResovler全局添加的）
//                .globalRequestParameters(getGlobalRequestParameters()); // 设置全局通用参数
    }

//    /**
//     * 项目通用参数，添加全局参数——登录认证token（若无可省略）
//     */
//    private List<RequestParameter> getGlobalRequestParameters() {
//        List<RequestParameter> parameters = new ArrayList<>();
//        parameters.add(new RequestParameterBuilder()
//                .name("token") // 参数名
//                .query(q -> q.model(m -> m.scalarModel(ScalarType.STRING))) //参数类型
//                .description("登录认证token") // 描述
//                .required(false) // 非必传
//                .in(ParameterType.HEADER) //请求头中的参数，其它类型可以点进ParameterType类中查看
//                .build());
//        return parameters;
//    }
}
