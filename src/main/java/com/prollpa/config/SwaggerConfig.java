package com.prollpa.config;


import org.springdoc.core.models.GroupedOpenApi;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.info.Info;

@Configuration
public class SwaggerConfig {

    // Define the general information about your API
    @Bean
    public OpenAPI customOpenAPI() {
        return new OpenAPI()
                .info(new Info()
                        .title("Prollpa API")
                        .version("1.0")
                        .description("This is the Prollpa API documentation."));
    }

    // Configure a specific group for the policy controller APIs
    @Bean
    public GroupedOpenApi policyGroup() {
        return GroupedOpenApi.builder()
                .group("com.prollpa") // This must match the name used in Swagger UI config
                .pathsToMatch("/**/**") // Match all paths or specific ones
                .build();
    }
}
