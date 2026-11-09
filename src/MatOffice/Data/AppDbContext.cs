using MatOffice.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;

namespace MatOffice.Data;

/// <summary>
/// The database (PostgreSQL). Tables are PascalCase and singular, the key is always <c>Id</c> (BIGINT), every record has CreateDate, CreateUserId, UpdateDate and UpdateUserId
/// (filled in <see cref="SaveChangesAsync(bool, CancellationToken)"/>), and a record with an <c>UpdateState</c> is deleted softly (0 = deleted, 1 = created, 2 = updated).
/// </summary>
public sealed class AppDbContext(DbContextOptions<AppDbContext> options, CurrentUser? current = null) : DbContext(options)
{
    public DbSet<User> Users => Set<User>();

    public DbSet<UserSession> Sessions => Set<UserSession>();

    public DbSet<Document> Documents => Set<Document>();

    public DbSet<DocumentMember> DocumentMembers => Set<DocumentMember>();

    public DbSet<DocumentVisit> DocumentVisits => Set<DocumentVisit>();

    public DbSet<DocumentVersion> DocumentVersions => Set<DocumentVersion>();

    public DbSet<HostClient> HostClients => Set<HostClient>();

    public DbSet<HostSession> HostSessions => Set<HostSession>();

    public DbSet<SystemSetting> SystemSettings => Set<SystemSetting>();

    protected override void OnModelCreating(ModelBuilder b)
    {
        b.Entity<User>(e =>
        {
            e.ToTable("User");
            e.HasIndex(x => x.Email).IsUnique();
            e.Property(x => x.Email).HasMaxLength(254);
            e.Property(x => x.DisplayName).HasMaxLength(120);
            e.Property(x => x.TimeZone).HasMaxLength(64);
            e.Property(x => x.Language).HasMaxLength(5);
            e.Property(x => x.Theme).HasMaxLength(10);
            e.Property(x => x.Scheme).HasMaxLength(10);
            e.Property(x => x.PaperSize).HasMaxLength(10);
            e.Property(x => x.ExternalSubject).HasMaxLength(200);
            e.HasIndex(x => x.ExternalSubject).IsUnique().HasFilter("\"ExternalSubject\" IS NOT NULL");
            e.HasQueryFilter(x => x.UpdateState != SoftEntity.Deleted);
        });

        b.Entity<UserSession>(e =>
        {
            e.ToTable("UserSession");
            e.HasIndex(x => x.Token).IsUnique();
            e.HasIndex(x => x.UserId);
            e.HasOne(x => x.User).WithMany().HasForeignKey(x => x.UserId).OnDelete(DeleteBehavior.Cascade);
            e.Property(x => x.UserAgent).HasMaxLength(300);
            e.Property(x => x.IpAddress).HasMaxLength(64);
            e.Property(x => x.IdentitySessionId).HasMaxLength(200);
            e.HasIndex(x => x.IdentitySessionId);
        });

        b.Entity<Document>(e =>
        {
            e.ToTable("Document");
            e.HasIndex(x => new { x.OwnerId, x.Kind });
            e.HasIndex(x => x.ShareToken).IsUnique();
            e.HasOne(x => x.Owner).WithMany().HasForeignKey(x => x.OwnerId).OnDelete(DeleteBehavior.Cascade);
            e.Property(x => x.Name).HasMaxLength(200);
            e.Property(x => x.Content).HasColumnType("jsonb");
            e.HasQueryFilter(x => x.UpdateState != SoftEntity.Deleted);
        });

        b.Entity<DocumentMember>(e =>
        {
            e.ToTable("DocumentMember");
            e.HasIndex(x => new { x.DocumentId, x.UserId }).IsUnique();
            e.HasIndex(x => x.UserId);
            e.HasOne(x => x.Document).WithMany(x => x.Members).HasForeignKey(x => x.DocumentId).OnDelete(DeleteBehavior.Cascade);
            e.HasOne(x => x.User).WithMany().HasForeignKey(x => x.UserId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<DocumentVisit>(e =>
        {
            e.ToTable("DocumentVisit");
            e.HasIndex(x => new { x.UserId, x.DocumentId }).IsUnique();
            e.HasIndex(x => new { x.UserId, x.VisitDate });
            e.HasOne(x => x.Document).WithMany().HasForeignKey(x => x.DocumentId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<DocumentVersion>(e =>
        {
            e.ToTable("DocumentVersion");
            e.HasIndex(x => new { x.DocumentId, x.CreateDate });
            e.HasOne(x => x.Document).WithMany().HasForeignKey(x => x.DocumentId).OnDelete(DeleteBehavior.Cascade);
            e.Property(x => x.Name).HasMaxLength(120);
            e.Property(x => x.Description).HasMaxLength(1000);
            e.Property(x => x.Content).HasColumnType("jsonb");
        });

        b.Entity<HostClient>(e =>
        {
            e.ToTable("HostClient");
            e.HasIndex(x => x.KeyHash).IsUnique();
            e.Property(x => x.Name).HasMaxLength(120);
            e.Property(x => x.ConsumerId).HasMaxLength(64);
            e.Property(x => x.ManagedBy).HasMaxLength(40);

            // one live binding per consumer (a removed host keeps its id for the record)
            e.HasIndex(x => x.ConsumerId).IsUnique().HasFilter("\"ConsumerId\" IS NOT NULL AND \"UpdateState\" <> 0");
            e.Property(x => x.KeyHash).HasMaxLength(64);
            e.Property(x => x.KeyPrefix).HasMaxLength(16);
            e.Property(x => x.AllowedUrlPrefix).HasMaxLength(300);
            e.Property(x => x.Note).HasMaxLength(1000);
            e.HasQueryFilter(x => x.UpdateState != SoftEntity.Deleted);
        });

        b.Entity<HostSession>(e =>
        {
            e.ToTable("HostSession");
            e.HasIndex(x => x.Token).IsUnique();
            e.HasIndex(x => x.ExpiresDate);
            e.HasOne(x => x.HostClient).WithMany().HasForeignKey(x => x.HostClientId).OnDelete(DeleteBehavior.Cascade);
            e.Property(x => x.FileId).HasMaxLength(200);
            e.Property(x => x.FileName).HasMaxLength(255);
            e.Property(x => x.Format).HasMaxLength(16);
            e.Property(x => x.ContentUrl).HasMaxLength(1000);
            e.Property(x => x.AccessToken).HasMaxLength(500);
            e.Property(x => x.UserRef).HasMaxLength(200);
            e.Property(x => x.UserName).HasMaxLength(120);
            e.Property(x => x.UserEmail).HasMaxLength(254);
            e.Property(x => x.ReturnUrl).HasMaxLength(1000);
            e.Property(x => x.LastError).HasMaxLength(1000);
            e.Property(x => x.HostVersion).HasMaxLength(200);
            e.Property(x => x.Content).HasColumnType("jsonb");
        });

        b.Entity<SystemSetting>(e =>
        {
            e.ToTable("SystemSetting");
            e.Property(x => x.PublicUrl).HasMaxLength(300);
        });
    }

    public override int SaveChanges(bool acceptAllChangesOnSuccess)
    {
        Stamp();
        return base.SaveChanges(acceptAllChangesOnSuccess);
    }

    public override Task<int> SaveChangesAsync(bool acceptAllChangesOnSuccess, CancellationToken cancellationToken = default)
    {
        Stamp();
        return base.SaveChangesAsync(acceptAllChangesOnSuccess, cancellationToken);
    }

    /// <summary>Fills the audit columns, and turns the delete of a soft record into UpdateState 0.</summary>
    private void Stamp()
    {
        var now = DateTime.UtcNow;
        var user = current?.Id;
        foreach (var entry in ChangeTracker.Entries<Entity>().ToList())
        {
            var entity = entry.Entity;
            if (entry.State == EntityState.Deleted && entity is SoftEntity soft)
            {
                entry.State = EntityState.Modified;
                soft.UpdateState = SoftEntity.Deleted;
            }

            switch (entry.State)
            {
                case EntityState.Added:
                    if (entity.CreateDate == default)
                    {
                        entity.CreateDate = now;
                    }

                    entity.CreateUserId ??= user;
                    entity.UpdateDate = now;
                    entity.UpdateUserId = user ?? entity.CreateUserId;
                    break;
                case EntityState.Modified:
                    entity.UpdateDate = now;
                    entity.UpdateUserId = user ?? entity.UpdateUserId;
                    if (entity is SoftEntity { UpdateState: not SoftEntity.Deleted } updated)
                    {
                        updated.UpdateState = SoftEntity.Updated;
                    }

                    break;
            }
        }
    }
}

/// <summary>For <c>dotnet ef migrations add</c>: no database is needed to create a migration.</summary>
public sealed class DesignTimeFactory : IDesignTimeDbContextFactory<AppDbContext>
{
    public AppDbContext CreateDbContext(string[] args)
        => new(new DbContextOptionsBuilder<AppDbContext>().UseNpgsql("Host=localhost;Database=matoffice;Username=matoffice;Password=matoffice").Options);
}
