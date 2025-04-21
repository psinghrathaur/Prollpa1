package com.prollpa.serviceImpl;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.prollpa.entity.Role;
import com.prollpa.entity.User;
import com.prollpa.entity.UserVscRole;
import com.prollpa.entity.VSC;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.UserVscRoleDto;
import com.prollpa.payload.UserVscRoleResponse;
import com.prollpa.repository.UserRepository;
import com.prollpa.repository.UserVSCRoleRepository;
import com.prollpa.service.RoleService;
import com.prollpa.service.UserVscRoleService;
import com.prollpa.service.VSCService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import com.prollpa.entity.VSC;
import com.prollpa.payload.RoleDto;

@Service
public class UserVscRoleServiceImpl implements UserVscRoleService {

    private static final Logger logger = LoggerFactory.getLogger(UserVscRoleServiceImpl.class);

    private final UserVSCRoleRepository userVSCRoleRepository;
    private final VSCService vscService;
    private final UserRepository userRepository;
    private final RoleService roleService;
    private final ModelMapper modelMapper;

    @Autowired
    public UserVscRoleServiceImpl(UserVSCRoleRepository userVSCRoleRepository, VSCService vscService,
                                  UserRepository userRepository, RoleService roleService, ModelMapper modelMapper) {
        this.userVSCRoleRepository = userVSCRoleRepository;
        this.vscService = vscService;
        this.userRepository = userRepository;
        this.roleService = roleService;
        this.modelMapper = modelMapper;
    }

    @Override
    public List<UserVscRoleDto> saveUserVscRoleList(List<UserVscRoleDto> userVscRoleList) {
        logger.info("Starting to save UserVscRole list");
        
        for (UserVscRoleDto dto : userVscRoleList) {
            if (dto.getVscId() == null || dto.getRoleId() == null || dto.getUserId() == null) {
                throw new ResourceNotFoundException("vscId ->"+dto.getVscId()+"and roleId->"+dto.getRoleId()+" userId: " + dto.getUserId()+" cannot be null");
            }
        }

        // Initialize list to avoid NullPointerException
        List<UserVscRole> userVscRoleEntities = new ArrayList<>();
        
        try {
            for (UserVscRoleDto userVscRoleDto : userVscRoleList) {
                logger.debug("Processing UserVscRoleDto: {}", userVscRoleDto);
                
                // Fetch user, VSC, and role from their respective services
                User user = userRepository.findById(userVscRoleDto.getUserId())
                        .orElseThrow(() -> new ResourceNotFoundException("User not found with this id: " + userVscRoleDto.getUserId()));
                VSC vsc = vscService.getVSCById(userVscRoleDto.getVscId());
                Role role = roleService.getRoleByRoleId(userVscRoleDto.getRoleId());
                
                if (user != null && vsc != null && role != null) {
                    UserVscRole userVscRole = new UserVscRole();
                    userVscRole.setRole(role);
                    userVscRole.setUser(user);
                    userVscRole.setVsc(vsc);
                    userVscRoleEntities.add(userVSCRoleRepository.save(userVscRole));
                    logger.debug("Saved UserVscRole: {}", userVscRole);
                }
            }
        } catch (Exception e) {
            logger.error("Error while saving UserVscRole list", e);
        }

        logger.info("Completed saving UserVscRole list");
        return mapUserVscRoleListtoDtoList(userVscRoleEntities);
    }

    @Override
    public List<UserVscRoleDto> getUserVscRoleByUserId(long userId) {
         List<UserVscRole> UserVscRolelist = userVSCRoleRepository.findByUser_userId(userId);
         if(UserVscRolelist.isEmpty()) {
        	 throw new ResourceNotFoundException("userVscRole not found with this userId"+userId);
         }
         return mapUserVscRoleListtoDtoList(UserVscRolelist);
    }

    @Override
    public List<UserVscRoleDto> getUserVscRoleByVscId(long vscId) {
        return mapUserVscRoleListtoDtoList(userVSCRoleRepository.findByVsc_vscId(vscId));
    }

    private List<UserVscRoleDto> mapUserVscRoleListtoDtoList(List<UserVscRole> userVscRoleList) {
        List<UserVscRoleDto> userVscRoleDtoList = new ArrayList<>();
        for (UserVscRole userVscRole : userVscRoleList) {
            userVscRoleDtoList.add(modelMapper.map(userVscRole, UserVscRoleDto.class));
        }
        return userVscRoleDtoList;
    }

    @Override
    public List<UserVscRoleResponse> getUserVscRoleByUserId2(List<UserVscRoleDto> userVscRoleList) {
        List<UserVscRoleResponse> userVscRoleResponsesList = new ArrayList<>();
        List<VSC> vscList = vscService.getVSCList(); // List of all VSC centers
        List<Role> roleList = roleService.getRoleList(); // List of all roles
        Map<Long, List<UserVscRoleDto>> groupedByVscId = userVscRoleList.stream()
                .collect(Collectors.groupingBy(UserVscRoleDto::getVscId));
        System.out.println(groupedByVscId);
        
        List<Map<String, Object>> result = new ArrayList<>();
        for (Map.Entry<Long, List<UserVscRoleDto>> entry : groupedByVscId.entrySet()) {
        	Long vscId = entry.getKey();
            List<UserVscRoleDto> vscRoleDtos = entry.getValue();
        	UserVscRoleResponse userVscRoleResponse = new UserVscRoleResponse();
            Map<String, Object> vscEntry = new HashMap<>();
            vscEntry.put("vsc_id", entry.getKey());
            String vscCenterName = vscList.stream()
                    .filter(vsc -> vsc.getVscId().equals(entry.getKey()))
                    .map(VSC::getVscCenterName)
                    .findFirst()
                    .orElse("Unknown VSC");
            userVscRoleResponse.setVscId(entry.getKey());
            userVscRoleResponse.setVscCenterName(vscCenterName);
            
            List<RoleDto> roles = vscRoleDtos.stream()
                    .map(dto -> {
                        Long roleId = dto.getRoleId();
                        String roleName = roleList.stream()
                                .filter(role -> Long.valueOf(role.getRoleId()).equals(roleId))
                                .map(Role::getRole)
                                .findFirst()
                                .orElse("Unknown Role");
                        
                        // Create and populate Role object
                        RoleDto role = new RoleDto();
                        role.setRoleId(roleId);
                        role.setRole(roleName);
                        return role;
                    })
                    .collect(Collectors.toList());
            

            userVscRoleResponse.setRole(roles);
            userVscRoleResponsesList.add(userVscRoleResponse);
        }

        // Return or use the response list
        return userVscRoleResponsesList;
//            
//
//            // Map roles
//            List<Map<String, Object>> roles = entry.getValue().stream()
//                    .map(dto -> {
//                        Map<String, Object> roleMap = new HashMap<>();
//                        roleMap.put("role_id", dto.getRoleId());
//                        // Add other role properties if needed
//                        return roleMap;
//                    })
//                    .collect(Collectors.toList());
//
//            vscEntry.put("role", roles);
//            result.add(vscEntry);
//        }
//        System.out.println(result);
//        // Sort by vsc_id
//        result.sort(Comparator.comparing(entry -> (Long) entry.get("vsc_id")));
//        System.out.println(result);
//
//        return null;//resul
//    
        
    }
}
