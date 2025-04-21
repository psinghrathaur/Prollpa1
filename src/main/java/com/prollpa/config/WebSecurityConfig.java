package com.prollpa.config;
//
//
//import org.springframework.context.annotation.Configuration;
//import org.springframework.ldap.core.LdapTemplate;
//import org.springframework.ldap.core.support.BaseLdapPathContextSource;
//import org.springframework.ldap.core.support.LdapContextSource;
//import org.springframework.security.config.annotation.authentication.builders.AuthenticationManagerBuilder;
//import org.springframework.security.config.annotation.authentication.configuration.AuthenticationConfiguration;
//import org.springframework.security.config.annotation.web.builders.HttpSecurity;
//import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
//import org.springframework.security.config.http.SessionCreationPolicy;
//import org.springframework.security.config.ldap.LdapBindAuthenticationManagerFactory;
//import org.springframework.security.core.userdetails.UserDetailsService;
//import org.springframework.security.authentication.AuthenticationManager;
//import org.springframework.security.config.Customizer;
//import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
//import org.springframework.security.crypto.password.PasswordEncoder;
//import org.springframework.security.web.SecurityFilterChain;
//import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;
//
//import com.prollpa.security.JwtAuthFilter;
//
//import org.springframework.context.annotation.Bean;
//import org.springframework.beans.factory.annotation.Autowired;
//
//
//@Configuration
//@EnableWebSecurity
//public class WebSecurityConfig {
//
//        @Bean
//	    public SecurityFilterChain securityFilterChain(HttpSecurity http) throws Exception {
//	        http
//	            .csrf().disable() // Disable CSRF for simplicity in JSON API
//	            .authorizeHttpRequests(auth -> auth
//	                .requestMatchers(
//	                    "/**", // Allow login endpoints
//	                    "/swagger-ui/**", // Allow Swagger UI
//	                    "/v3/api-docs/**" ,
//	                    "/test1"
//	                  // Allow OpenAPI specs
//	                ).permitAll()
//	                .anyRequest().authenticated() // Secure all other endpoints
//	            )
//	            .httpBasic().disable() // Disable HTTP Basic authentication
//	            .formLogin().disable()
//	            .addFilterBefore(new JwtAuthFilter(), UsernamePasswordAuthenticationFilter.class)
//                .build();;/*
//	            .rememberMe()
//                .key("uniqueAndSecret") // A unique key for cookie encryption
//                .tokenValiditySeconds(86400) // Expiry time of cookies (1 day)
//            .and()
//            .sessionManagement()
//                .sessionCreationPolicy(SessionCreationPolicy.IF_REQUIRED) // Create session only when needed
//                .invalidSessionUrl("/auth/login") // Redirect to login if session is invalid
//            ;// Disable form-based authentication*/
//
//	        return http.build();
//	    }
//  @Bean
//  public LdapTemplate ldapTemplate() {
//	  return new LdapTemplate(contentSource());
//  }
//  
//  @Bean
//  public LdapContextSource contentSource() {
//	  LdapContextSource ldapContentSource=new LdapContextSource();
//	  ldapContentSource.setUrl("ldap://localhost:10389");
//      ldapContentSource.setUserDn("uid=admin,ou=system");
//	  ldapContentSource.setPassword("secret");
//	  return ldapContentSource;
//  }
//
//  @Bean 
//  AuthenticationManager authManager(BaseLdapPathContextSource source) {
//	  LdapBindAuthenticationManagerFactory factory=new LdapBindAuthenticationManagerFactory(source);
//	  factory.setUserDnPatterns("uid={0},ou=users,ou=system");
//	  return factory.createAuthenticationManager();
//  }
//  @Bean
//  public AuthenticationManager authenticationManager(AuthenticationConfiguration config) throws Exception {
//      return config.getAuthenticationManager();
//  }
//  @Bean
//  public PasswordEncoder passwordEncoder() {
//      return new BCryptPasswordEncoder(); // Password encoding
//  }
//
//}
//




import com.prollpa.security.JwtAuthFilter;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.ldap.core.LdapTemplate;
import org.springframework.ldap.core.support.BaseLdapPathContextSource;
import org.springframework.ldap.core.support.LdapContextSource;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.config.annotation.authentication.configuration.AuthenticationConfiguration;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.config.ldap.LdapBindAuthenticationManagerFactory;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;

@Configuration
@EnableWebSecurity
public class WebSecurityConfig {

    @Bean
    public SecurityFilterChain securityFilterChain(HttpSecurity http, JwtAuthFilter jwtAuthFilter) throws Exception {
        http
            .csrf().disable().cors().and()
            .authorizeHttpRequests(auth -> auth
                .requestMatchers(
                		"/auth/**",          // your authentication endpoints
                        "/swagger-ui/**",    // swagger UI
                        "/v3/api-docs/**",   // swagger docs
                        "/swagger-ui.html",
                        "/v3/api-docs/com.prollpa",
                        "/swagger-ui/oauth2-redirect.html"
//                        "/excel/**"// access to swagger html
                ).permitAll()
                .anyRequest().authenticated()
            )
            .sessionManagement()
            .sessionCreationPolicy(SessionCreationPolicy.STATELESS)
            .and()
            .addFilterBefore(jwtAuthFilter, UsernamePasswordAuthenticationFilter.class);

        return http.build();
    }

    @Bean
    public LdapTemplate ldapTemplate(LdapContextSource contextSource) {
        return new LdapTemplate(contextSource);
    }

    @Bean
    public LdapContextSource contextSource() {
        LdapContextSource contextSource = new LdapContextSource();
        contextSource.setUrl("ldap://localhost:10389");
        contextSource.setUserDn("uid=admin,ou=system");
        contextSource.setPassword("secret");
        return contextSource;
    }

    @Bean
    public AuthenticationManager authenticationManager(BaseLdapPathContextSource contextSource) {
        LdapBindAuthenticationManagerFactory factory = new LdapBindAuthenticationManagerFactory(contextSource);
        factory.setUserDnPatterns("uid={0},ou=users,ou=system");
        return factory.createAuthenticationManager();
    }

    @Bean
    public PasswordEncoder passwordEncoder() {
        return new BCryptPasswordEncoder();
    }
}


