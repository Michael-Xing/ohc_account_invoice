create table if not exists t_omd_file_name_mapping
(
    id            varchar(36)                                not null
        primary key,
    standard_name varchar(255) default ''::character varying not null,
    name_ja       varchar(255),
    name_zh       varchar(255),
    name_en       varchar(255)
);

alter table t_omd_file_name_mapping
    owner to postgres;


create table if not exists t_omd_production_area_mapping
(
    production_area varchar not null
        constraint t_omd_production_area_map_pk
            primary key,
    country         varchar,
    address         varchar,
    phone           varchar
);

alter table t_omd_production_area_mapping
    owner to postgres;

create table if not exists t_omd_project_file_chunk_info_ja
(
    id                      varchar(36),
    project_file_id         varchar(36)                                        not null,
    slice_summary           text,
    slice_content           text,
    create_time             timestamp with time zone default CURRENT_TIMESTAMP not null,
    update_time             timestamp with time zone default CURRENT_TIMESTAMP not null,
    deleted                 boolean,
    slice_content_embedding vector,
    slice_summary_embedding vector,
    slice_index             integer,
    slice_section           varchar(255)
);

alter table t_omd_project_file_chunk_info_ja
    owner to postgres;

create table if not exists t_omd_project_file_chunk_info_zh
(
    id                      varchar(36),
    project_file_id         varchar(36)                                        not null,
    slice_summary           text,
    slice_content           text,
    create_time             timestamp with time zone default CURRENT_TIMESTAMP not null,
    update_time             timestamp with time zone default CURRENT_TIMESTAMP not null,
    deleted                 boolean,
    slice_content_embedding vector,
    slice_summary_embedding vector,
    slice_index             integer,
    slice_section           varchar(255)
);

alter table t_omd_project_file_chunk_info_zh
    owner to postgres;

create table if not exists t_omd_project_file_info
(
    id             varchar(36),
    name           varchar(255)                                       not null,
    project_number varchar(50)                                        not null,
    stage          varchar(20)                                        not null,
    version        varchar(20)                                        not null,
    keyword_info   jsonb,
    create_time    timestamp with time zone default CURRENT_TIMESTAMP not null,
    update_time    timestamp with time zone default CURRENT_TIMESTAMP not null,
    deleted        boolean,
    file_type      varchar(255)                                       not null,
    file_number    varchar(100),
    need_ocr       boolean,
    short_name     varchar(255),
    ocr_status     varchar(20),
    stage_2th      varchar(50),
    file_path      varchar(512),
    short_name_id  varchar(20)
);

alter table t_omd_project_file_info
    owner to postgres;


create table if not exists t_omd_test_file_chunk_info_ja
(
    id                      varchar(36),
    test_file_id            varchar(36),
    slice_summary           text,
    slice_content           text,
    slice_summary_embedding vector,
    slice_content_embedding vector,
    create_time             timestamp with time zone,
    update_time             timestamp with time zone,
    deleted                 boolean,
    slice_index             integer,
    slice_section           varchar(255)
);

alter table t_omd_test_file_chunk_info_ja
    owner to postgres;



create table if not exists t_omd_test_file_chunk_info_zh
(
    id                      varchar(36),
    test_file_id            varchar(36),
    slice_summary           text,
    slice_content           text,
    slice_summary_embedding vector,
    slice_content_embedding vector,
    create_time             timestamp with time zone,
    update_time             timestamp with time zone,
    deleted                 boolean,
    slice_index             integer,
    slice_section           varchar(255)
);

alter table t_omd_test_file_chunk_info_zh
    owner to postgres;


create table if not exists t_omd_test_file_info
(
    id            varchar(36),
    name          varchar(255),
    file_path     varchar(512),
    product_name  varchar(255),
    standard_name varchar(255),
    version       varchar(50),
    test_number   varchar(100),
    test_name     varchar(255),
    keyword_info  jsonb,
    need_ocr      boolean,
    ocr_status    varchar(20),
    create_time   timestamp with time zone,
    update_time   timestamp with time zone,
    deleted       boolean
);

alter table t_omd_test_file_info
    owner to postgres;

create table if not exists t_omd_test_name_map
(
    name_ja      varchar,
    name_zh      varchar,
    product_name varchar,
    test_number  varchar(50)
);

alter table t_omd_test_name_map
    owner to postgres;


create table if not exists t_omd_sharepoint_sync_files
(
    id              INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    sharepoint_id   varchar(255)                                                            not null,
    sharepoint_path varchar(1000)                                                           not null,
    local_path      varchar(1000)                                                           not null,
    file_name       varchar(255)                                                            not null,
    file_size       integer                                                                 not null,
    last_modified   timestamp                                                               not null,
    etag            varchar(255),
    checksum        varchar(128),
    sync_status     varchar(50),
    error_message   text,
    created_at      timestamp,
    updated_at      timestamp
);

alter table t_omd_sharepoint_sync_files
    owner to postgres;

create unique index ix_t_omd_sharepoint_sync_files_sharepoint_ids
    on t_omd_sharepoint_sync_files (sharepoint_id);

create index ix_t_omd_sharepoint_sync_files_sharepoint_paths
    on t_omd_sharepoint_sync_files (sharepoint_path);


-- auto-generated definition
create table if not exists t_omd_sharepoint_sync_logs
(
    id              INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    operation       varchar(100)                                                           not null,
    sharepoint_path varchar(1000),
    local_path      varchar(1000),
    status          varchar(50)                                                            not null,
    message         text,
    duration_ms     integer,
    file_size       integer,
    created_at      timestamp
);

alter table t_omd_sharepoint_sync_logs
    owner to postgres;


